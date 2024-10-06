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

    <div v-if="sheets">
      <TabView v-model:activeIndex="activeTabIndex">
        <TabPanel v-for="(sheet, sheetName) in sheets" :key="sheetName" :header="sheetName">
          <div v-if="selectedFormula !== null" class="formula-bar">
            <span>Selected Formula: </span>
            <input 
              v-model="selectedFormula" 
              @keyup.enter="updateFormula"
              @blur="updateFormula"
              class="formula-input"
            />
          </div>
          
          <h2>{{ sheetName }}</h2>
          <DataTable :value="sheet.data" editMode="cell" @cell-edit-complete="onCellEditComplete" class="editable-cells-table">
            <Column v-for="col in sheet.column_names" :key="col" :field="col">
              <template #header>{{ col }}</template>
              <template #body="slotProps">
                <div @click="onCellClick(slotProps.index, col, sheetName)">
                  {{ getCellDisplayValue(slotProps.data[col]) }}
                </div>
              </template>
              <template #editor="{ data, field }">
                <InputText v-model="data[field]" />
              </template>
            </Column>
          </DataTable>
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
      const sheetName = Object.keys(sheets.value)[activeTabIndex.value]
      const sheet = sheets.value[sheetName]
      const rowIndex = cellCoord.row.index
      const colIndex = cellCoord.column.index
      const cellValue = sheet.data[rowIndex][sheet.column_names[colIndex]]
      done(cellValue)
    })

    parser.on('callRangeValue', function(startCellCoord, endCellCoord, done) {
      const sheetName = Object.keys(sheets.value)[activeTabIndex.value]
      const sheet = sheets.value[sheetName]
      const fragment = []
      for (let row = startCellCoord.row.index; row <= endCellCoord.row.index; row++) {
        const rowData = []
        for (let col = startCellCoord.column.index; col <= endCellCoord.column.index; col++) {
          rowData.push(sheet.data[row][sheet.column_names[col]])
        }
        fragment.push(rowData)
      }
      done(fragment)
    })

    const onCellEditComplete = (event) => {
      const { newValue, index, field } = event
      const sheetName = Object.keys(sheets.value)[activeTabIndex.value]
      sheets.value[sheetName].data[index][field] = newValue
    }

    const onCellClick = (rowIndex, column, sheetName) => {
      const cellValue = sheets.value[sheetName].data[rowIndex][column]
      selectedFormula.value = typeof cellValue === 'string' && cellValue.startsWith('=') ? cellValue : null
      selectedCell.value = selectedFormula.value ? { rowIndex, column, sheetName } : null
    }

    const updateFormula = () => {
      if (selectedCell.value) {
        const { rowIndex, column, sheetName } = selectedCell.value
        sheets.value[sheetName].data[rowIndex][column] = selectedFormula.value
      }
    }

    return {
      activeTabIndex,
      sheets,
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