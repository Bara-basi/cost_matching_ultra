<script setup>
defineProps({ direction: { type: String, default: 'right' } })
</script>

<template>
  <svg class="ui-arrow" viewBox="0 0 16 16" aria-hidden="true" focusable="false">
    <path v-if="direction === 'left'" d="m10 3-5 5 5 5" />
    <path v-else-if="direction === 'down'" d="m3 6 5 5 5-5" />
    <path v-else d="m6 3 5 5-5 5" />
  </svg>
</template>

<style scoped>
.ui-arrow{display:block;width:13px;height:13px;flex:none;vertical-align:middle;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round;fill:none}
</style>
