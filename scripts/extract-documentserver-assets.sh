#!/usr/bin/env bash
# extract-documentserver-assets.sh — 从 OnlyOffice DocumentServer 镜像导出静态资源到 public/
#
# 用途（升级 OnlyOffice 静态资源）
#   本脚本是本项目更新 OnlyOffice SDK 静态资源（fonts / sdkjs / web-apps / sdkjs-plugins）
#   的推荐方式
# 升级后必须复查的接入层 patch（SDK 全量替换后不会自动保留）
#   字体
#     - sdkjs/common/AllFonts.js 中的 __custom_font_registry__、自定义 fonts/{id} 产物
#     - README.zh.md「字体配置」；fonts/ttf-to-catalog-font.mjs 等工具脚本是否需一并拷贝
#     - editor-manager.ts：installIframeProxies 对 AllFonts.js / libfont 的原生 XHR 绕过
#   批注 / 修订（依赖 sdkjs/word 内部 API，版本差异大时需手调）
#     - core/editor-manager.ts：pluginMethod_AddComment、asc_* 修订栈 / report 回填、
#       refreshCommentsFromSdk / refreshRevisionsFromSdk、Word 内容同步回调
#     - feature/comments.ts、feature/revisions.ts
#     - scripts/test-comment-revision-apis.mjs 跑一遍回归
#   文档加载 / x2t 转换（否则易出现 Editor.bin 为空、批注修订读不到）
#     - internal/editor/utils.ts：getX2tConvertFormats / getX2tExportFormats（formatTo 须为 CANVAS 类型）
#     - internal/editor/server.ts：loadDocument 传入 formatFrom / formatTo
#     - internal/editor/types.ts：AvsFileType 枚举是否与新 x2t 一致
#
# Usage:
#   ./scripts/extract-documentserver-assets.sh
#   ./scripts/extract-documentserver-assets.sh public/packages/onlyoffice/9.3.0
#   ./scripts/extract-documentserver-assets.sh --no-pull
#   IMAGE=docker-0.unsee.tech/onlyoffice/documentserver:9.3.0 ./scripts/extract-documentserver-assets.sh

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IMAGE="${IMAGE:-docker-0.unsee.tech/onlyoffice/documentserver:9.3.0}"
CONTAINER_SRC="/var/www/onlyoffice/documentserver"
OUT_DIR="${ROOT}/public/9.3.0-backup"
DO_PULL=1

DIRS=(fonts sdkjs web-apps sdkjs-plugins)

usage() {
  sed -n '2,36p' "$0" | sed 's/^# \{0,1\}//'
  exit "${1:-0}"
}

log() { printf '→ %s\n' "$*"; }
die() { printf '✗ %s\n' "$*" >&2; exit 1; }

parse_args() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      -h|--help) usage 0 ;;
      --no-pull) DO_PULL=0; shift ;;
      --image)
        [[ $# -ge 2 ]] || die "--image 需要参数"
        IMAGE="$2"
        shift 2
        ;;
      --*)
        die "未知参数: $1（使用 --help 查看用法）"
        ;;
      *)
        OUT_DIR="$1"
        shift
        ;;
    esac
  done

  case "$OUT_DIR" in
    /*) ;;
    *) OUT_DIR="${ROOT}/${OUT_DIR}" ;;
  esac
}

require_docker() {
  if ! command -v docker >/dev/null 2>&1; then
    die "未找到 docker 命令，请先安装 Docker"
  fi
  if ! docker info >/dev/null 2>&1; then
    die "Docker 未运行，请先启动 Docker Desktop"
  fi
}

ensure_image() {
  if docker image inspect "$IMAGE" >/dev/null 2>&1; then
    log "使用本地镜像: ${IMAGE}"
    return
  fi
  if [[ "$DO_PULL" -eq 0 ]]; then
    die "本地不存在镜像 ${IMAGE}，去掉 --no-pull 或先 docker pull"
  fi
  log "拉取镜像: ${IMAGE}"
  docker pull "$IMAGE"
}

remove_target_dir() {
  local target="$1"
  [[ -e "$target" ]] || return 0
  chmod -R u+rwX "$target" 2>/dev/null || chmod -R a+rwX "$target" 2>/dev/null || true
  rm -rf "$target"
  if [[ -e "$target" ]]; then
    die "无法删除旧目录: ${target}（请手动: chmod -R u+w '${target}' && rm -rf '${target}'）"
  fi
}

extract_all_assets() {
  local name dest

  log "输出目录: ${OUT_DIR}"
  for name in "${DIRS[@]}"; do
    remove_target_dir "${OUT_DIR}/${name}"
  done
  mkdir -p "$OUT_DIR"

  log "运行 documentserver-generate-allfonts.sh false（与 Dockerfile 相同）…"
  log "导出 fonts / sdkjs / web-apps / sdkjs-plugins …"

  # 生成脚本写 AllFonts.js -> sdkjs/common/，字体二进制 -> fonts/，须在同一个容器内再打包
  if ! docker run --rm --entrypoint sh "$IMAGE" -c \
    "documentserver-generate-allfonts.sh false >&2 && tar -C \"${CONTAINER_SRC}\" -cf - fonts sdkjs web-apps sdkjs-plugins" \
    | tar -xf - -C "$OUT_DIR"; then
    die "生成或 tar 导出失败"
  fi

  for name in "${DIRS[@]}"; do
    dest="${OUT_DIR}/${name}"
    [[ -d "$dest" ]] || die "缺少目录: ${dest}"
    chmod -R u+rwX "$dest" 2>/dev/null || true
  done

  local api_tpl="${OUT_DIR}/web-apps/apps/api/documents/api.js.tpl"
  local api_js="${OUT_DIR}/web-apps/apps/api/documents/api.js"
  if [[ -f "$api_tpl" ]]; then
    log "复制 api.js.tpl -> api.js（与 Dockerfile 相同）"
    cp "$api_tpl" "$api_js"
  fi

  [[ -f "${OUT_DIR}/sdkjs/common/AllFonts.js" ]] \
    || die "缺少 sdkjs/common/AllFonts.js（请确认 generate 脚本已执行）"

  local font_count
  font_count="$(find "${OUT_DIR}/fonts" -type f 2>/dev/null | wc -l | tr -d ' ')"
  if [[ "$font_count" -eq 0 ]]; then
    die "fonts 目录为空"
  fi
}

print_post_upgrade_reminder() {
  cat <<'EOF'

⚠  静态资源已替换，请手动复查以下接入层 patch（不会随 tar 导出自动保留）：

  字体
    · 合并 AllFonts.js 中的 __custom_font_registry__ 与自定义 fonts/{id}
    · 见 README.zh.md「字体配置」

  批注 / 修订
    · src/components/onlyoffice-web-comp/core/editor-manager.ts
    · feature/comments.ts、feature/revisions.ts
    · 建议：node scripts/test-comment-revision-apis.mjs

  文档加载 / x2t
    · internal/editor/utils.ts（getX2tConvertFormats，formatTo 使用 CANVAS 类型）
    · internal/editor/server.ts、util/x2t.ts
    · 若升级了 x2t WASM，同步 public/.../x2t/

  站点路径
    · const/index.ts 中 STATIC_RESOURCE / NEXT_PUBLIC_APP_ROOT 与导出目录一致

EOF
}

print_summary() {
  echo ""
  echo "✓ 导出完成"
  echo "  镜像: ${IMAGE}"
  echo "  目标: ${OUT_DIR}"
  echo "  AllFonts.js: ${OUT_DIR}/sdkjs/common/AllFonts.js"
  echo ""
  printf "  %-16s %s\n" "目录" "大小"
  for name in "${DIRS[@]}"; do
    du -sh "${OUT_DIR}/${name}" | awk -v n="$name" '{printf "  %-16s %s\n", n, $1}'
  done
  echo ""
  for name in "${DIRS[@]}"; do
    count="$(find "${OUT_DIR}/${name}" -type f 2>/dev/null | wc -l | tr -d ' ')"
    printf "  %s: %s 个文件\n" "$name" "$count"
  done
  print_post_upgrade_reminder
}

main() {
  parse_args "$@"
  require_docker
  ensure_image
  extract_all_assets
  print_summary
}

main "$@"
