import { act, createElement, type ReactElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'

vi.mock('@tiptap/react', async () => {
  const react = await import('react')
  return {
    NodeViewWrapper: ({
      children,
      className,
    }: {
      children: import('react').ReactNode
      className?: string
    }) => react.createElement('div', { className }, children),
    NodeViewContent: () => react.createElement('code'),
  }
})

import { CodeBlockView } from '../src/renderer/editor/CodeBlockView'

const mountedRoots: Array<{ root: Root; container: HTMLElement }> = []

function mount(element: ReactElement): { root: Root; container: HTMLElement } {
  const container = document.createElement('div')
  document.body.appendChild(container)
  const root = createRoot(container)
  act(() => root.render(element))
  mountedRoots.push({ root, container })
  return { root, container }
}

function unmount(root: Root): void {
  const mounted = mountedRoots.find((entry) => entry.root === root)
  act(() => root.unmount())
  mounted?.container.remove()
  if (mounted) mountedRoots.splice(mountedRoots.indexOf(mounted), 1)
}

afterEach(() => {
  for (const { root, container } of mountedRoots.splice(0)) {
    act(() => root.unmount())
    container.remove()
  }
  vi.useRealTimers()
  vi.restoreAllMocks()
})

describe('CodeBlockView teardown', () => {
  it('clears copy feedback timeout when its node view unmounts', async () => {
    vi.useFakeTimers()
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { writeText: vi.fn(async () => {}) },
    })
    const clearTimeout = vi.spyOn(window, 'clearTimeout')
    const { root, container } = mount(
      createElement(CodeBlockView, {
        node: { attrs: { language: null }, textContent: 'const answer = 42' },
        updateAttributes: vi.fn(),
        editor: { isEditable: true },
      } as never),
    )

    await act(async () => {
      container.querySelector<HTMLButtonElement>('.md-codeblock-copy')!.click()
      await Promise.resolve()
    })
    unmount(root)

    expect(clearTimeout).toHaveBeenCalled()
  })
})
