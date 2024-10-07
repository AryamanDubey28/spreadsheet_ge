<template>
  <div class="excel-table">
    <div class="file-upload">
      <input type="file" @change="handleFileUpload" accept=".xlsx,.csv" />
    </div>

    <div v-if="error" class="error-message">
      {{ error }}
    </div>

    <div v-if="loading" class="loading-message">
      Loading and processing file...
    </div>

    <div v-if="sheets" class="sheet-container">
      <div class="sheet-content">
        <div v-if="selectedCell" class="formula-bar">
          <span>Selected Cell: </span>
          <input 
            v-model="editingFormula"
            @keyup.enter="updateFormula"
            @blur="updateFormula"
            class="formula-input"
          />
          <button @click="cancelFormulaEdit" class="cancel-button">
            <span class="cancel-icon">×</span>
          </button>
        </div>
        
        <h2>{{ currentSheetName }}</h2>
        <DataTable :value="currentSheetData" editMode="cell" @cell-edit-complete="onCellEditComplete" class="editable-cells-table" :scrollable="true" scrollHeight="flex">
          <Column v-for="col in currentColumns" :key="col" :field="col">
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
        <TabPanel v-for="sheetName in sheetNames" :key="sheetName" :header="sheetName">
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
    const selectedCell = ref(null)
    const editingFormula = ref('')

    const sheetNames = computed(() => sheets.value ? Object.keys(sheets.value) : [])

    const currentSheetName = computed(() => sheetNames.value[activeTabIndex.value] || '')

    const currentSheet = computed(() => sheets.value && currentSheetName.value ? sheets.value[currentSheetName.value] : null)

    const currentSheetData = computed(() => currentSheet.value ? currentSheet.value.data : [])

    const currentColumns = computed(() => currentSheet.value ? currentSheet.value.column_names : [])

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
        setupParser()
      } catch (error) {
        throw new Error('Error uploading file to server: ' + error.message)
      }
    }

    const setupParser = () => {
      parser.on('callCellValue', function(cellCoord, done) {
        const rowIndex = cellCoord.row.index - 1
        const colIndex = cellCoord.column.index
        
        if (rowIndex < 0 || rowIndex >= currentSheetData.value.length ||
            colIndex < 0 || colIndex >= currentColumns.value.length) {
          return done(null);
        }

        let cellValue = currentSheetData.value[rowIndex][currentColumns.value[colIndex]]
        if (!isNaN(cellValue) && cellValue !== '') {
          cellValue = Number(cellValue)
        }
        console.log(`At row ${rowIndex + 1} col ${currentColumns.value[colIndex]}, cellvalue = ${cellValue}`)
        done(cellValue)
      })

      parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
        const fragment = []
        for (let row = startCellCoord.row.index - 1; row <= endCellCoord.row.index - 1; row++) {
          const rowData = []
          for (let col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
            if (row >= 0 && row < currentSheetData.value.length &&
                col >= 0 && col < currentColumns.value.length) {
              rowData.push(currentSheetData.value[row][currentColumns.value[col]])
            } else {
              rowData.push(null)
            }
          }
          fragment.push(rowData)
        }
        done(fragment)
      })
    }

    const getCellDisplayValue = (value) => {
      if (typeof value === 'string' && value.startsWith('=')) {
        return evaluateFormula(value)
      }
      return value
    }

    const evaluateFormula = (formula) => {
      const result = parser.parse(formula.slice(1)) 
      console.log(`formula = ${formula.slice(1)} result = ${JSON.stringify(result)}`)
      return result.error ? result.error : result.result
    }

    const onCellEditComplete = (event) => {
      const { newValue, index, field } = event
      currentSheetData.value[index][field] = newValue
    }

    const onCellClick = (rowIndex, column) => {
      selectedCell.value = { rowIndex: rowIndex + 1, column }
      const cellValue = currentSheetData.value[rowIndex][column]
      editingFormula.value = cellValue
    }

    const updateFormula = () => {
      if (selectedCell.value) {
        const { rowIndex, column } = selectedCell.value
        currentSheetData.value[rowIndex - 1][column] = editingFormula.value
      }
      selectedCell.value = null
      editingFormula.value = ''
    }

    const cancelFormulaEdit = () => {
      selectedCell.value = null
      editingFormula.value = ''
    }

    return {
      sheets,
      sheetNames,
      activeTabIndex,
      currentSheetName,
      currentSheetData,
      currentColumns,
      selectedCell,
      editingFormula,
      error,
      loading,
      handleFileUpload,
      onCellEditComplete,
      getCellDisplayValue,
      onCellClick,
      updateFormula,
      cancelFormulaEdit
    }
  },
})
</script>

<style scoped>
.excel-table {
  display: flex;
  flex-direction: column;
  height: 100vh;
  padding: 1rem;
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

.sheet-container {
  display: flex;
  flex-direction: column;
  flex-grow: 1;
  overflow: hidden;
}

.sheet-content {
  flex-grow: 1;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.formula-bar {
  background-color: #f0f0f0;
  padding: 0.5rem;
  margin-bottom: 0.5rem;
  border: 1px solid #ccc;
  border-radius: 4px;
  display: flex;
  align-items: center;
}

.formula-input {
  flex-grow: 1;
  margin-left: 0.5rem;
  margin-right: 0.5rem;
  padding: 0.25rem;
  border: 1px solid #ccc;
  border-radius: 4px;
}

.cancel-button {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 1.2rem;
  padding: 0 0.5rem;
}

.cancel-icon {
  color: #888;
}

.cancel-icon:hover {
  color: #333;
}

.editable-cells-table {
  flex-grow: 1;
}

:deep(.p-datatable-wrapper) {
  max-height: unset !important;
}

.sheet-tabs {
  margin-top: 1rem;
}
</style>