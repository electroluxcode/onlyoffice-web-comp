/**
 * onlyoffice-comp 公共 API
 *
 * | 层        | 目录      | 职责                              |
 * |-----------|-----------|-----------------------------------|
 * | const     | const/    | 常量、静态资源、文件类型映射      |
 * | store     | store/    | 文档/语言等跨页面状态             |
 * | util      | util/     | SDK 初始化、x2t 转换、下载工具    |
 * | core      | core/     | EditorManager、事件总线、业务门面 |
 * | feature   | feature/  | 批注、修订                        |
 * | internal  | internal/ | mock server / socket（不对外）    |
 */

export * from "./const";
export * from "./store";
export * from "./util";
export * from "./core";
export * from "./feature";

export type * from "./type/word-api";
export type * from "./type/sdk-internal";
export type { OfficeTheme, PluginMode } from "./internal/editor/types";
