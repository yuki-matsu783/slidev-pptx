<script setup lang="ts">
import { ref } from 'vue'

const count = ref(0)
</script>

<template>
  <button
    class="px-4 py-2 rounded bg-teal-600 text-white hover:bg-teal-500"
    @click="count++"
  >
    count: {{ count }}
  </button>
</template>
