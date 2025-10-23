// Vue 3 wrapper component for CustomTable (vanilla JS)
// Usage with CDN:
// <script src="https://unpkg.com/vue@3/dist/vue.global.prod.js"></script>
// <script type="module">
//   import { createCustomTableComponent } from './lib/vue-custom-table.js';
//   const CustomTableVue = createCustomTableComponent(Vue);
//   const app = Vue.createApp({
//     components: { CustomTableVue },
//     template: `<CustomTableVue ref=\"grid\" :options="{ rows: 4, cols: 4 }" />`,
//     mounted() {
//       // Access underlying instance
//       const table = this.$refs.grid.getInstance();
//       table.setCellStyle(0, 0, { background: '#ffeeaa' });
//     }
//   });
//   app.mount('#app');
// </script>

import { CustomTable } from './custom-table.js';

export function createCustomTableComponent(Vue) {
  const { defineComponent, h, ref, onMounted, onBeforeUnmount, watch, nextTick } = Vue;

  return defineComponent({
    name: 'CustomTableVue',
    props: {
      // Initial internal model (rows, cols, data, styles, merges). If also passing 'sheet', 'sheet' wins.
      model: { type: Object, default: null },
      // Spreadsheet JSON (commercial.json-like). If provided, it will be loaded via fromJSON.
      sheet: { type: Object, default: null },
      // Options for new table if no model/sheet provided (e.g., { rows: 2, cols: 3 })
      options: { type: Object, default: () => ({}) },
      // When true, emits 'change' with the current internal model on user edits (debounced)
      emitChange: { type: Boolean, default: false },
      // Debounce time for change emits (ms)
      debounce: { type: Number, default: 100 }
    },
    emits: ['ready', 'change'],
    setup(props, { emit, expose }) {
      const hostRef = ref(null);
      let table = null;
      let changeTimer = null;

      const mountTable = () => {
        if (!hostRef.value) return;
        table = new CustomTable(hostRef.value, props.options || {});
        if (props.sheet) {
          table.fromJSON(props.sheet);
        } else if (props.model) {
          table.setModel(props.model);
        }
        if (props.emitChange) attachChangeListeners();
        emit('ready', table);
      };

      const destroyTable = () => {
        if (table && typeof table.destroy === 'function') {
          table.destroy();
        }
        table = null;
      };

      const attachChangeListeners = () => {
        // Generic input/click listeners on host to detect edits and style changes
        const el = hostRef.value;
        if (!el) return;
        const scheduleEmit = () => {
          if (!props.emitChange) return;
          if (changeTimer) clearTimeout(changeTimer);
          changeTimer = setTimeout(() => {
            if (table) emit('change', table.getModel());
          }, props.debounce);
        };
        el.addEventListener('input', scheduleEmit);
        el.addEventListener('click', (e) => {
          // Only debounce for toolbar/style actions and add/remove; ignore pure selection focus
          const target = e.target;
          if (target && (target.closest('.ct-toolbar') || target.closest('.add-col') || target.closest('.add-row') || target.closest('.remove-col') || target.closest('.row-head'))) {
            scheduleEmit();
          }
        });
      };

      onMounted(() => {
        nextTick(mountTable);
      });

      onBeforeUnmount(() => {
        destroyTable();
      });

      // React to prop changes (replace full table state)
      watch(() => props.sheet, (val) => {
        if (!table) return;
        if (val && typeof table.fromJSON === 'function') table.fromJSON(val);
      }, { deep: true });

      watch(() => props.model, (val) => {
        if (!table) return;
        if (val && typeof table.setModel === 'function') table.setModel(val);
      }, { deep: true });

      // expose imperative API to parent via template ref
      expose({
        getInstance: () => table,
        toJSON: () => table?.toJSON?.(),
        toSpreadsheetJSON: () => table?.toSpreadsheetJSON?.(),
        exportToExcel: (filename) => table?.exportToExcel?.(filename),
        exportToCSV: (filename) => table?.exportToCSV?.(filename),
        getModel: () => table?.getModel?.(),
        setModel: (m) => table?.setModel?.(m),
        fromJSON: (j) => table?.fromJSON?.(j)
      });

      return () => h('div', { ref: hostRef, class: 'ctable-vue-host' });
    }
  });
}

// If Vue is available globally, attach a global factory for convenience
if (typeof window !== 'undefined' && window.Vue && !window.CustomTableVue) {
  try {
    window.CustomTableVue = createCustomTableComponent(window.Vue);
  } catch {}
}
