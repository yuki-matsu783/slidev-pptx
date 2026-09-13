<script setup lang="ts">
// 試作: PPT 部品の「文字を編集」ボタンと、中身の Markdown を編集する小窓。
// ボタンは Slidev のドラッグ枠（#drag-control-container）の中に出す。枠の外を押すと選択が外れるため。
// components/ に置かないのは、Slidev に部品として自動登録させないため。
import type { PptTag } from '../composables/usePptEdit'
import { nextTick, ref } from 'vue'
import { usePptEdit } from '../composables/usePptEdit'

const props = defineProps<{ tag: PptTag, drag: string, dragging: boolean }>()
const emit = defineEmits<{ open: [] }>()

const { editing, saving, draft, error, open, save, cancel } = usePptEdit(props, props.tag)
const area = ref<HTMLTextAreaElement | null>(null)

async function onOpen() {
  emit('open')
  await open()
  await nextTick()
  area.value?.focus()
}

function onKeydown(ev: KeyboardEvent) {
  // 日本語の変換を確定する Enter では何もしない
  if (ev.isComposing) return
  if (ev.key === 'Escape') {
    ev.preventDefault()
    cancel()
  }
  else if (ev.key === 'Enter' && (ev.metaKey || ev.ctrlKey)) {
    ev.preventDefault()
    save()
  }
}
</script>

<template>
  <Teleport v-if="props.dragging && !editing" defer to="#drag-control-container">
    <button type="button" class="ppt-edit-btn" @pointerdown.stop @click.stop="onOpen">
      文字を編集
    </button>
  </Teleport>
  <Teleport v-if="editing" to="body">
    <div class="ppt-edit-backdrop" @pointerdown.self="cancel">
      <div class="ppt-edit-panel" role="dialog" :aria-label="`${props.tag} の中身を編集`">
        <div class="ppt-edit-head">
          <code>&lt;{{ props.tag }} drag="{{ props.drag }}"&gt;</code> の中身（Markdown）
        </div>
        <p v-if="error" class="ppt-edit-error">
          {{ error }}
        </p>
        <textarea
          ref="area"
          v-model="draft"
          class="ppt-edit-area"
          rows="8"
          spellcheck="false"
          :disabled="!!error"
          @keydown="onKeydown"
        />
        <div class="ppt-edit-actions">
          <span class="ppt-edit-hint">⌘ / Ctrl + Enter で保存、Esc で閉じる</span>
          <button type="button" class="ppt-edit-cancel" @click="cancel">
            キャンセル
          </button>
          <button type="button" class="ppt-edit-save" :disabled="saving || !!error" @click="save">
            {{ saving ? '保存中…' : '保存' }}
          </button>
        </div>
      </div>
    </div>
  </Teleport>
</template>

<style scoped>
.ppt-edit-btn {
  position: absolute;
  top: -34px;
  right: 0;
  z-index: 200;
  padding: 3px 10px;
  font: 500 13px/1.4 system-ui, sans-serif;
  color: #fff;
  background: #0d7a69;
  border: 0;
  border-radius: 4px;
  cursor: pointer;
  white-space: nowrap;
}
.ppt-edit-btn:focus-visible {
  outline: 2px solid #0d7a69;
  outline-offset: 2px;
}
.ppt-edit-backdrop {
  position: fixed;
  inset: 0;
  z-index: 10000;
  display: grid;
  place-items: center;
  padding: 16px;
  background: rgb(15 23 27 / 0.45);
}
.ppt-edit-panel {
  width: min(640px, 100%);
  display: grid;
  gap: 10px;
  padding: 16px;
  color: #16232a;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 12px 40px rgb(0 0 0 / 0.25);
  font: 14px/1.6 system-ui, sans-serif;
}
.ppt-edit-head code {
  font-family: ui-monospace, Menlo, Consolas, monospace;
  font-size: 13px;
}
.ppt-edit-error {
  margin: 0;
  padding: 6px 10px;
  color: #7a3b00;
  background: #fbf0de;
  border-radius: 4px;
}
.ppt-edit-area {
  width: 100%;
  box-sizing: border-box;
  padding: 10px;
  font: 14px/1.6 ui-monospace, Menlo, Consolas, monospace;
  color: inherit;
  background: #f4f7f7;
  border: 1px solid #c9d4d7;
  border-radius: 4px;
  resize: vertical;
}
.ppt-edit-area:focus {
  outline: 2px solid #0d7a69;
  outline-offset: -1px;
}
.ppt-edit-actions {
  display: flex;
  align-items: center;
  gap: 8px;
}
.ppt-edit-hint {
  margin-right: auto;
  font-size: 12px;
  color: #56666e;
}
.ppt-edit-actions button {
  padding: 5px 14px;
  font: inherit;
  border-radius: 4px;
  cursor: pointer;
}
.ppt-edit-cancel {
  color: inherit;
  background: transparent;
  border: 1px solid #c9d4d7;
}
.ppt-edit-save {
  color: #fff;
  background: #0d7a69;
  border: 1px solid #0d7a69;
}
.ppt-edit-save:disabled {
  opacity: 0.5;
  cursor: default;
}
</style>
