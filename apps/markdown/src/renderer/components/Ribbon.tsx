import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { Editor } from '@tiptap/core'
import { useEditorState } from '@tiptap/react'
import {
  Dropdown,
  RibbonCollapseButton,
  RibbonExpandButton,
  useDismissablePopover,
  useRibbonCollapse,
} from '@genoffice/ui'
import { useI18n } from '../i18n/locale'
import type { StringKey } from '../i18n/locale'
import { uiOp, type BlockType, type ListKind, type StylableMark } from '../editor/ops'
import {
  IconBullets,
  IconHr,
  IconInlineCode,
  IconLink,
  IconNumbered,
  IconOutlineView,
  IconPicture,
  IconProperties,
  IconRedo,
  IconSave,
  IconSearch,
  IconSourceCode,
  IconSpellcheck,
  IconTable,
  IconTaskList,
  IconUndo,
} from './icons'

interface Props {
  editor: Editor | null
  disabled: boolean
  dirty: boolean
  onSave: () => void
  onSaveAs: () => void
  onFind: () => void
  autoSave: boolean
  onToggleAutoSave: (on: boolean) => void
  imageEnabled: boolean
  onInsertImage: () => void
  frontmatterOpen: boolean
  onToggleFrontmatter: () => void
  /** the source view replaces the document canvas, so editor-shaped commands go dead */
  /** this PR's markdown source view — NOT the .txt/.json `sourceMode` */
  sourceViewOpen: boolean
  onToggleSource: () => void
  outlineOpen: boolean
  onToggleOutline: () => void
  hasOutline: boolean
  spellcheck: boolean
  onToggleSpellcheck: () => void
  /**
   * A .txt/.json is edited as source, so every button that formats a block
   * document is hidden rather than left there to act on a document that does
   * not exist. Save, find and autosave still apply.
   */
  sourceMode?: boolean
}

type BlockStyle = 'paragraph' | 'h1' | 'h2' | 'h3' | 'h4' | 'h5' | 'h6' | 'quote' | 'codeBlock'

const STYLE_LABEL: Record<BlockStyle, StringKey> = {
  paragraph: 'styleParagraph',
  h1: 'styleH1',
  h2: 'styleH2',
  h3: 'styleH3',
  h4: 'styleH4',
  h5: 'styleH5',
  h6: 'styleH6',
  quote: 'styleQuote',
  codeBlock: 'styleCodeBlock',
}

function applyBlockStyle(editor: Editor, style: BlockStyle): void {
  const type: BlockType =
    style === 'quote'
      ? 'blockquote'
      : style === 'paragraph' || style === 'codeBlock'
        ? style
        : 'heading'
  const level = type === 'heading' ? Number(style.slice(1)) : undefined
  uiOp(editor, { op: 'setBlockType', target: 'selection', type, ...(level ? { level } : {}) })
}

const toggleStyle = (editor: Editor | null, style: StylableMark) =>
  editor && uiOp(editor, { op: 'setStyle', target: 'selection', style, mode: 'toggle' })

const toggleList = (editor: Editor | null, list: ListKind) =>
  editor && uiOp(editor, { op: 'toggleList', target: 'selection', list })

function IconBtn({
  title,
  active,
  disabled,
  onClick,
  children,
}: {
  title: string
  active?: boolean
  disabled?: boolean
  onClick: () => void
  children: ReactNode
}) {
  return (
    <button
      type="button"
      className={`rb-btn${active ? ' active' : ''}`}
      data-tip={title}
      aria-label={title}
      disabled={disabled}
      onMouseDown={(e) => e.preventDefault()}
      onClick={onClick}
    >
      {children}
    </button>
  )
}

export function Ribbon({
  editor,
  disabled,
  dirty,
  onSave,
  onSaveAs,
  onFind,
  autoSave,
  onToggleAutoSave,
  imageEnabled,
  onInsertImage,
  frontmatterOpen,
  onToggleFrontmatter,
  sourceViewOpen,
  onToggleSource,
  outlineOpen,
  onToggleOutline,
  hasOutline,
  spellcheck,
  onToggleSpellcheck,
  sourceMode = false,
}: Props) {
  const { t } = useI18n()
  const collapse = useRibbonCollapse('mdapp.ribbonCollapsed', {
    collapse: t('ribbonCollapse'),
    expand: t('ribbonExpand'),
  })
  const [linkOpen, setLinkOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const linkInputRef = useRef<HTMLInputElement>(null)
  const linkAnchorRef = useRef<HTMLSpanElement>(null)

  const state = useEditorState({
    editor,
    selector: ({ editor: e }) => {
      if (!e) return null
      const style: BlockStyle = e.isActive('codeBlock')
        ? 'codeBlock'
        : e.isActive('blockquote')
          ? 'quote'
          : e.isActive('heading')
            ? (`h${e.getAttributes('heading').level}` as BlockStyle)
            : 'paragraph'
      return {
        style,
        bold: e.isActive('bold'),
        italic: e.isActive('italic'),
        strike: e.isActive('strike'),
        code: e.isActive('code'),
        link: e.isActive('link'),
        bullet: e.isActive('bulletList'),
        ordered: e.isActive('orderedList'),
        task: e.isActive('taskList'),
        canUndo: e.can().undo(),
        canRedo: e.can().redo(),
      }
    },
  })

  useEffect(() => {
    if (linkOpen) linkInputRef.current?.focus()
  }, [linkOpen])

  useDismissablePopover(linkOpen, () => setLinkOpen(false), {
    inside: () => [linkAnchorRef.current],
  })

  // In the source view the canvas is a textarea: formatting acts on a selection
  // the user cannot see, and undo/focus would yank them out of the pane, so the
  // editor-shaped commands stand down. Saving and the source toggle stay live.
  const off = disabled || sourceMode || sourceViewOpen || !editor || !state

  const openLink = () => {
    if (!editor) return
    setLinkUrl(String(editor.getAttributes('link').href ?? ''))
    setLinkOpen((v) => !v)
  }

  const applyLink = () => {
    if (!editor) return
    uiOp(editor, { op: 'setLink', target: 'selection', href: linkUrl.trim() || null })
    setLinkOpen(false)
  }

  // 20px inline-row rendering, same as the docs toolbar these icons come from
  // (pinned stroke paints 1.5px at this size per the suite-wide icon rules)
  const ICON = 20

  return (
    <div className={`ribbon ${collapse.rootClass}`} ref={collapse.rootRef}>
      {/* quick-access row above the toolbar (save / undo / redo / autosave), same as the docs QAT row */}
      <div className="ribbon-tabs">
        <button
          type="button"
          className="qa-btn"
          data-tip={t('save')}
          aria-label={t('save')}
          disabled={disabled || !dirty}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onSave}
        >
          <IconSave size={16} />
        </button>
        <button
          type="button"
          className="qa-btn qa-save-as"
          data-tip={t('saveAs')}
          aria-label={t('saveAs')}
          disabled={off}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onSaveAs}
        >
          {t('saveAs')}
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('undo')}
          aria-label={t('undo')}
          disabled={off || !state?.canUndo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => editor?.chain().focus().undo().run()}
        >
          <IconUndo size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('redo')}
          aria-label={t('redo')}
          disabled={off || !state?.canRedo}
          onMouseDown={(e) => e.preventDefault()}
          onClick={() => editor?.chain().focus().redo().run()}
        >
          <IconRedo size={16} />
        </button>
        <button
          type="button"
          className="qa-btn"
          data-tip={t('findTip')}
          aria-label={t('findTip')}
          disabled={off}
          onMouseDown={(e) => e.preventDefault()}
          onClick={onFind}
        >
          <IconSearch size={16} />
        </button>
        <label className={`autosave-toggle${autoSave ? ' on' : ''}`} data-tip={t('autoSaveTip')}>
          <span className="autosave-knob" />
          <span className="autosave-text">{t('autoSave')}</span>
          <input
            type="checkbox"
            checked={autoSave}
            onChange={(e) => onToggleAutoSave(e.target.checked)}
          />
        </label>
        <RibbonExpandButton state={collapse} label={t('ribbonExpand')} />
      </div>

      <div className="ribbon-body" data-ribbon-body="">
        {!sourceMode && (
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <Dropdown
                className="rb-style"
                value={state?.style ?? 'paragraph'}
                disabled={off}
                options={(Object.keys(STYLE_LABEL) as BlockStyle[]).map((s) => ({
                  value: s,
                  label: t(STYLE_LABEL[s]),
                }))}
                onPick={(s) => editor && applyBlockStyle(editor, s)}
              />
            </div>
          </div>
        )}

        {!sourceMode && <div className="rb-sep" />}

        {!sourceMode && (
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <IconBtn
                title={t('bold')}
                active={state?.bold}
                disabled={off}
                onClick={() => toggleStyle(editor, 'bold')}
              >
                <b>B</b>
              </IconBtn>
              <IconBtn
                title={t('italic')}
                active={state?.italic}
                disabled={off}
                onClick={() => toggleStyle(editor, 'italic')}
              >
                <i>I</i>
              </IconBtn>
              <IconBtn
                title={t('strike')}
                active={state?.strike}
                disabled={off}
                onClick={() => toggleStyle(editor, 'strike')}
              >
                <s>ab</s>
              </IconBtn>
              <IconBtn
                title={t('inlineCode')}
                active={state?.code}
                disabled={off}
                onClick={() => toggleStyle(editor, 'code')}
              >
                <IconInlineCode size={ICON} />
              </IconBtn>
              <span className="rb-link-anchor" ref={linkAnchorRef}>
                <IconBtn title={t('link')} active={state?.link} disabled={off} onClick={openLink}>
                  <IconLink size={ICON} />
                </IconBtn>
                {linkOpen && (
                  <span className="rb-link-pop" onMouseDown={(e) => e.stopPropagation()}>
                    <input
                      ref={linkInputRef}
                      value={linkUrl}
                      placeholder={t('linkPlaceholder')}
                      onChange={(e) => setLinkUrl(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && linkUrl.trim()) applyLink()
                        if (e.key === 'Escape') setLinkOpen(false)
                      }}
                    />
                    <button type="button" disabled={!linkUrl.trim()} onClick={applyLink}>
                      {t('linkApply')}
                    </button>
                  </span>
                )}
              </span>
            </div>
          </div>
        )}

        {!sourceMode && <div className="rb-sep" />}

        {!sourceMode && (
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <IconBtn
                title={t('bulletList')}
                active={state?.bullet}
                disabled={off}
                onClick={() => toggleList(editor, 'bullet')}
              >
                <IconBullets size={ICON} />
              </IconBtn>
              <IconBtn
                title={t('orderedList')}
                active={state?.ordered}
                disabled={off}
                onClick={() => toggleList(editor, 'ordered')}
              >
                <IconNumbered size={ICON} />
              </IconBtn>
              <IconBtn
                title={t('taskList')}
                active={state?.task}
                disabled={off}
                onClick={() => toggleList(editor, 'task')}
              >
                <IconTaskList size={ICON} />
              </IconBtn>
            </div>
          </div>
        )}

        {!sourceMode && <div className="rb-sep" />}

        {!sourceMode && (
          <div className="ribbon-group">
            <div className="ribbon-group-items">
              <IconBtn
                title={t('insertTable')}
                disabled={off}
                onClick={() => editor && uiOp(editor, { op: 'insertTable', after: 'selection' })}
              >
                <IconTable size={ICON} />
              </IconBtn>
              <IconBtn
                title={t('insertImage')}
                disabled={off || !imageEnabled}
                onClick={onInsertImage}
              >
                <IconPicture size={ICON} />
              </IconBtn>
              <IconBtn
                title={t('insertHr')}
                disabled={off}
                onClick={() =>
                  editor && uiOp(editor, { op: 'insertHorizontalRule', after: 'selection' })
                }
              >
                <IconHr size={ICON} />
              </IconBtn>
            </div>
          </div>
        )}

        <div className="rb-spacer" />

        <div className="ribbon-group">
          <div className="ribbon-group-items">
            {/*
              The markdown source view is the PR's own mode, distinct from
              `sourceMode` (a .txt/.json edited as source). Both hide the
              block-formatting buttons: there is no block document to format in
              either case.
            */}
            {/* A .txt/.json is already source text, so there is nothing for
                this toggle to switch to; it is a markdown control. */}
            {!sourceMode && (
              <IconBtn
                title={t('sourceView')}
                active={sourceViewOpen}
                disabled={disabled}
                onClick={onToggleSource}
              >
                <IconSourceCode size={ICON} />
              </IconBtn>
            )}
            {!sourceMode && (
              <IconBtn
                title={t('fmProperties')}
                active={frontmatterOpen}
                disabled={disabled}
                onClick={onToggleFrontmatter}
              >
                <IconProperties size={ICON} />
              </IconBtn>
            )}
            {!sourceMode && (
              <IconBtn
                title={t('outline')}
                active={outlineOpen}
                disabled={disabled || (!hasOutline && !outlineOpen)}
                onClick={onToggleOutline}
              >
                <IconOutlineView size={ICON} />
              </IconBtn>
            )}
            <IconBtn
              title={t('spellcheck')}
              active={spellcheck}
              disabled={disabled}
              onClick={onToggleSpellcheck}
            >
              <IconSpellcheck size={ICON} />
            </IconBtn>
          </div>
        </div>
      </div>
      <RibbonCollapseButton state={collapse} label={t('ribbonCollapse')} />
    </div>
  )
}
