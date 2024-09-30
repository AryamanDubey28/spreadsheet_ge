<template>
  <div class="excel-table">
    <h2>Excel-like Table</h2>
    <DataTable :value="displayData" editMode="cell" @cell-edit-complete="onCellEditComplete" class="editable-cells-table">
      <Column v-for="col in columns" :key="col" :field="col">
        <template #header>{{ col }}</template>
        <template #body="slotProps">
          {{ getCellDisplayValue(slotProps.data[col]) }}
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
  { A: '=MIN(A1:A7)', B: '=COUNT(B1:B7)', C: '=CONCATENATE("Sum: ", SUM(C1:C3))' },
  { A: 15, B: '=SQRT(A9)', C: '=ROUNDDOWN(B9, 1)' },
  { A: '=ABS(A1-A9)', B: '=MOD(A10, 3)', C: '=AND(A10>5, B10<2)' }
])

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

    return {
      columns,
      displayData,
      onCellEditComplete,
      getCellDisplayValue,
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
</style>