<template>
  <div class="excel-table">
    <div class="file-upload">
      <input type="file" @change="handleFileUpload" accept=".xlsx,.csv" />
    </div>

    <div v-if="error" class="error-message">
      {{ error }}
    </div>

    <div v-if="loading" class="loading-message">
      Loading and uploading file...
    </div>

    <div v-if="sheets" class="sheet-container">
      <div class="sheet-content">
        <div v-if="selectedFormula !== null" class="formula-bar">
          <span>Selected Formula: </span>
          <input 
            v-model="selectedFormula" 
            @keyup.enter="updateFormula"
            @blur="updateFormula"
            class="formula-input"
          />
        </div>
        
        <h2>{{ currentSheetName }}</h2>
        <DataTable :value="currentSheet.data" editMode="cell" @cell-edit-complete="onCellEditComplete" class="editable-cells-table">
          <Column v-for="col in currentSheet.column_names" :key="col" :field="col">
            <template #header>{{ col }}</template>
            <template #body="slotProps">
              <div @click="onCellClick(slotProps.index, col)">
                {{ getCellDisplayValue(slotProps.data[col]) }}
              </div>
            </template>
            <template #editor="{ data, field }">
              <InputText v-model="data[field]" />
            </template>
          </Column>
        </DataTable>
      </div>
      
      <TabView v-model:activeIndex="activeTabIndex" class="sheet-tabs">
        <TabPanel v-for="(sheet, sheetName) in sheets" :key="sheetName" :header="sheetName">
        </TabPanel>
      </TabView>
    </div>
  </div>
</template>

<script>
import { defineComponent, ref, computed } from 'vue'
import { Parser } from 'hot-formula-parser'
import axios from 'axios'

export default defineComponent({
  name: 'ExcelTable',
  setup() {
    const parser = new Parser()
    const error = ref(null)
    const loading = ref(false)
    const sheets = ref(null)
    const activeTabIndex = ref(0)
    const selectedFormula = ref(null)
    const selectedCell = ref(null)

    const currentSheetName = computed(() => {
      if (!sheets.value) return ''
      return Object.keys(sheets.value)[activeTabIndex.value]
    })

    const currentSheet = computed(() => {
      if (!sheets.value || !currentSheetName.value) return { data: [], column_names: [] }
      return sheets.value[currentSheetName.value]
    })

    const handleFileUpload = async (event) => {
      const file = event.target.files[0]
      if (!file) return

      error.value = null
      loading.value = true

      try {
        await uploadFile(file)
      } catch (err) {
        console.error('Error processing file:', err)
        error.value = "Error processing and uploading file. Please try again."
      } finally {
        loading.value = false
      }
    }

    const uploadFile = async (file) => {
      const formData = new FormData()
      formData.append('file', file)

      try {
        const response = await axios.post('http://localhost:8000/upload-excel/', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        })
        sheets.value = response.data.sheets
        activeTabIndex.value = 0
      } catch (error) {
        throw new Error('Error uploading file to server: ' + error.message)
      }
    }

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
      const cellValue = currentSheet.value.data[rowIndex][currentSheet.value.column_names[colIndex]]
      done(cellValue)
    })

    parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
      const fragment = []
      for (let row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
        const rowData = []
        for (let col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
          rowData.push(currentSheet.value.data[row][currentSheet.value.column_names[col]])
        }
        fragment.push(rowData)
      }
      done(fragment)
    })

    const onCellEditComplete = (event) => {
      const { newValue, index, field } = event
      currentSheet.value.data[index][field] = newValue
    }

    const onCellClick = (rowIndex, column) => {
      const cellValue = currentSheet.value.data[rowIndex][column]
      selectedFormula.value = typeof cellValue === 'string' && cellValue.startsWith('=') ? cellValue : null
      selectedCell.value = selectedFormula.value ? { rowIndex, column } : null
    }

    const updateFormula = () => {
      if (selectedCell.value) {
        const { rowIndex, column } = selectedCell.value
        currentSheet.value.data[rowIndex][column] = selectedFormula.value
      }
    }

    return {
      activeTabIndex,
      sheets,
      currentSheetName,
      currentSheet,
      onCellEditComplete,
      getCellDisplayValue,
      selectedFormula,
      onCellClick,
      updateFormula,
      handleFileUpload,
      error,
      loading,
    }
  },
})
</script>

<style scoped>
.excel-table {
  margin-top: 2rem;
  display: flex;
  flex-direction: column;
  height: calc(100vh - 2rem); /* Adjust based on your layout */
}
.sheet-container {
  display: flex;
  flex-direction: column;
  flex-grow: 1;
  overflow: hidden;
}
.sheet-content {
  flex-grow: 1;
  overflow-y: auto;
  padding-bottom: 1rem;
}
.sheet-tabs {
  flex-shrink: 0;
}
.editable-cells-table ::v-deep(.p-datatable-wrapper) {
  overflow-y: auto;
  max-height: calc(100vh - 200px); /* Adjust based on your layout */
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
  display: flex;
  align-items: center;
}
.formula-input {
  flex-grow: 1;
  margin-left: 10px;
  padding: 5px;
  border: 1px solid #ccc;
  border-radius: 4px;
}
.file-upload {
  margin-bottom: 1rem;
}
.error-message {
  color: red;
  margin-bottom: 1rem;
}
.loading-message {
  color: blue;
  margin-bottom: 1rem;
}
</style>