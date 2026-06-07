'use client'

import { useState, useEffect } from 'react'
import { usePathname, useRouter } from 'next/navigation'
import { Layout, Menu, Button } from 'antd'
import type { MenuProps } from 'antd'
import {
  FileExcelOutlined,
  FileWordOutlined,
  FilePptOutlined,
  AppstoreOutlined,
  FolderOutlined,
  GithubOutlined,
  MenuOutlined,
  CloseOutlined,
  CodeOutlined,
} from '@ant-design/icons'
import './styles.css'

const { Sider } = Layout

const menuItems: MenuProps['items'] = [
  {
    key: '/excel',
    icon: <FileExcelOutlined />,
    label: 'Excel',
    children: [
      {
        key: '/excel/base',
        icon: <FolderOutlined />,
        label: 'Base',
      },
    ],
  },
  {
    key: '/docs',
    icon: <FileWordOutlined />,
    label: 'Docs',
    children: [
      {
        key: '/docs/base',
        icon: <FolderOutlined />,
        label: 'Base',
      },
    ],
  },
  {
    key: '/ppt',
    icon: <FilePptOutlined />,
    label: 'PPT',
    children: [
      {
        key: '/ppt/base',
        icon: <FolderOutlined />,
        label: 'Base',
      },
    ],
  },
  {
    key: '/multi',
    icon: <AppstoreOutlined />,
    label: 'Multi Instance',
    children: [
      {
        key: '/multi/base',
        icon: <FolderOutlined />,
        label: 'Base',
      },
    ],
  },
  {
    key: '/editor',
    icon: <CodeOutlined />,
    label: 'Editor',
    children: [
      {
        key: '/editor',
        icon: <FolderOutlined />,
        label: 'Web Office',
      },
    ],
  },
]

interface StudioLayoutProps {
  children: React.ReactNode
}

export default function StudioLayout({ children }: StudioLayoutProps) {
  const pathname = usePathname()
  const router = useRouter()
  const [openKeys, setOpenKeys] = useState<string[]>([])
  const [sidebarOpen, setSidebarOpen] = useState(false)

  // 根据路径自动展开对应的菜单并选中菜单项
  useEffect(() => {
    if (!pathname) return

    // 所有可能的路径映射
    const routeMap: Record<string, { selected: string; parent: string }> = {
      '/excel/base': { selected: '/excel/base', parent: '/excel' },
      '/docs/base': { selected: '/docs/base', parent: '/docs' },
      '/ppt/base': { selected: '/ppt/base', parent: '/ppt' },
      '/multi/base': { selected: '/multi/base', parent: '/multi' },
      '/editor': { selected: '/editor', parent: '/editor' },
    }

    // 精确匹配路径
    const match = routeMap[pathname]
    if (match) {
      setOpenKeys([match.parent])
    }
  }, [pathname])

  // 获取当前选中的菜单项
  const getSelectedKeys = (): string[] => {
    if (!pathname) return []
    return [pathname]
  }

  const handleMenuClick: MenuProps['onClick'] = ({ key }) => {
    router.push(key as string)
    // 移动端点击菜单项后自动关闭侧边栏
    if (window.innerWidth <= 768) {
      setSidebarOpen(false)
    }
  }

  const handleOpenChange: MenuProps['onOpenChange'] = (keys) => {
    setOpenKeys(keys)
  }

  const toggleSidebar = () => {
    setSidebarOpen(!sidebarOpen)
  }

  const closeSidebar = () => {
    setSidebarOpen(false)
  }

  return (
    <Layout className="studio-layout">
      {/* 移动端遮罩层 */}
      {sidebarOpen && (
        <div className="sidebar-overlay" onClick={closeSidebar} />
      )}

      {/* 移动端菜单按钮 */}
      <Button
        type="text"
        icon={<MenuOutlined />}
        className="mobile-menu-button"
        onClick={toggleSidebar}
        aria-label="切换菜单"
      />

      <Sider
        width={240}
        className={`studio-sidebar ${sidebarOpen ? 'open' : ''}`}
        theme="light"
      >
        <div className="sidebar-header">
          <div className="sidebar-header-content">
            <h2 className="sidebar-title">OnlyOffice Studio</h2>
            <Button
              type="text"
              icon={<CloseOutlined />}
              className="mobile-close-button"
              onClick={closeSidebar}
              aria-label="关闭菜单"
            />
          </div>
        </div>
        
        <Menu
          mode="inline"
          selectedKeys={getSelectedKeys()}
          openKeys={openKeys}
          items={menuItems}
          onClick={handleMenuClick}
          onOpenChange={handleOpenChange}
          className="sidebar-menu"
        />

        <div className="sidebar-footer">
          <div className="sidebar-info">
            <p>Powered by OnlyOffice</p>
            <p className="version">v1.0.0</p>
            <a
              href="https://github.com/electroluxcode/mvp-onlyoffice"
              target="_blank"
              rel="noopener noreferrer"
              className="github-link"
              title="查看 GitHub 仓库"
            >
              <GithubOutlined /> GitHub
            </a>
          </div>
        </div>
      </Sider>

      <Layout className="studio-main-content">
        {children}
      </Layout>
    </Layout>
  )
}
