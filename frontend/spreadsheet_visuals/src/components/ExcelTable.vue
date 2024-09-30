<template>
  <div class="excel-table">
    <div v-if="selectedFormula" class="formula-bar">
      Selected Formula: {{ selectedFormula }}
    </div>
    <h2>Excel-like Table</h2>
    <DataTable :value="displayData" editMode="cell" @cell-edit-complete="onCellEditComplete" class="editable-cells-table">
      <Column v-for="col in columns" :key="col" :field="col">
        <template #header>{{ col }}</template>
        <template #body="slotProps">
          <div @click="onCellClick(slotProps.data.index, col)">
            {{ getCellDisplayValue(slotProps.data[col]) }}
          </div>
        </template>
        <template #editor="{ data, field }">
          <InputText v-model="data[field]" />
        </template>
      </Column>
    </DataTable>
  </div>
</template>

<script>
import { defineComponent, ref, computed, reactive } from 'vue'
import { Parser } from 'hot-formula-parser'

export default defineComponent({
  name: 'ExcelTable',
  setup() {
    const parser = new Parser()
    const columns = ref(['A', 'B', 'C'])
    const tableData = reactive([
      { A: 1, B: 2, C: '=A1+B1' },
      { A: 3, B: 4, C: '=A2*B2' },
      { A: 5, B: 6, C: '=SUM(A1:B2)' },
      { A: 10, B: 20, C: '=AVERAGE(A1:A4)' },
      { A: 7, B: 8, C: '=MAX(B1:B5)' },
      { A: '=A4-A3', B: '=B4/B3', C: '=IF(A6>B6, "A is larger", "B is larger or equal")' },
      { A: 2, B: 3, C: '=POWER(A7,B7)' },
    ])

    const selectedFormula = ref(null)

    const getCellDisplayValue = (value) => {
      if (typeof value === 'string' && value.startsWith('=')) {
        return evaluateFormula(value)
      }
      return value
    }

    const evaluateFormula = (formula) => {
      const result = parser.parse(formula.slice(1))
      return result.error ? result.error : result.result
    }

    parser.on('callCellValue', function(cellCoord, done) {
      const rowIndex = cellCoord.row.index
      const colIndex = cellCoord.column.index
      const cellValue = tableData[rowIndex][columns.value[colIndex]]
      done(cellValue)
    })

    parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
      const fragment = []
      for (let row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
        const rowData = []
        for (let col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
          rowData.push(tableData[row][columns.value[col]])
        }
        fragment.push(rowData)
      }
      done(fragment)
    })

    const displayData = computed(() => {
      return tableData.map((row, index) => ({
        ...row,
        index
      }))
    })

    const onCellEditComplete = (event) => {
      const { newValue, index, field } = event
      tableData[index][field] = newValue
    }

    const onCellClick = (rowIndex, column) => {
      if (column === 'C') {
        const cellValue = tableData[rowIndex][column]
        if (typeof cellValue === 'string' && cellValue.startsWith('=')) {
          selectedFormula.value = cellValue
        } else {
          selectedFormula.value = null
        }
      } else {
        selectedFormula.value = null
      }
    }

    return {
      columns,
      displayData,
      onCellEditComplete,
      getCellDisplayValue,
      selectedFormula,
      onCellClick,
    }
  },
})
</script>

<style scoped>
.excel-table {
  margin-top: 2rem;
}
.editable-cells-table ::v-deep(.p-cell-editing) {
  padding-top: 0 !important;
  padding-bottom: 0 !important;
}
.formula-bar {
  background-color: #f0f0f0;
  padding: 10px;
  margin-bottom: 10px;
  border: 1px solid #ccc;
  border-radius: 4px;
}
</style>