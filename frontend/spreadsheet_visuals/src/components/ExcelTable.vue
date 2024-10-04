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

    <div v-if="selectedFormula !== null" class="formula-bar">
      <span>Selected Formula: </span>
      <input 
        v-model="selectedFormula" 
        @keyup.enter="updateFormula"
        @blur="updateFormula"
        class="formula-input"
      />
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
import axios from 'axios'
import * as XLSX from 'xlsx'

export default defineComponent({
  name: 'ExcelTable',
  setup() {
    const columns = ref([])
    const tableData = reactive([])
    const parser = new Parser()
    const error = ref(null)
    const loading = ref(false)

    const handleFileUpload = async (event) => {
      const file = event.target.files[0]
      if (!file) return

      error.value = null
      loading.value = true

      try {
        await readAndUploadFile(file)
      } catch (err) {
        console.error('Error processing file:', err)
        error.value = "Error processing and uploading file. Please try again."
      } finally {
        loading.value = false
      }
    }

    const readAndUploadFile = async (file) => {
      // Read and parse the file
      const fileData = await readFile(file)
      
      // Update frontend data
      columns.value = fileData.columns
      tableData.splice(0, tableData.length, ...fileData.data)

      // Upload to backend
      await uploadFile(file)
    }

    const readFile = (file) => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = (e) => {
          try {
            const result_data = new Uint8Array(e.target.result)
            const workbook = XLSX.read(result_data, { type: 'array' })
            const firstSheetName = workbook.SheetNames[0]
            const worksheet = workbook.Sheets[firstSheetName]
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
            
            const columns = jsonData[0]
            const data = jsonData.slice(1).map(row => {
              return columns.reduce((obj, key, index) => {
                obj[key] = row[index]
                return obj
              }, {})
            })

            resolve({ columns, data })
          } catch (error) {
            reject(error)
          }
        }
        reader.onerror = reject
        reader.readAsArrayBuffer(file)
      })
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
        console.log('File uploaded successfully:', response.data)
      } catch (error) {
        throw new Error('Error uploading file to server: ' + error.message)
      }
    }

    const displayData = computed(() => {
      return tableData.map((row, index) => ({
        ...row,
        index
      }))
    })

    const selectedFormula = ref(null)
    const selectedCell = ref(null)

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

    const onCellEditComplete = (event) => {
      const { newValue, index, field } = event
      tableData[index][field] = newValue
    }

    const onCellClick = (rowIndex, column) => {
      const cellValue = tableData[rowIndex][column]
      selectedFormula.value = typeof cellValue === 'string' && cellValue.startsWith('=') ? cellValue : null
      selectedCell.value = selectedFormula.value ? { rowIndex, column } : null
    }

    const updateFormula = () => {
      if (selectedCell.value) {
        const { rowIndex, column } = selectedCell.value
        tableData[rowIndex][column] = selectedFormula.value
      }
    }

    return {
      columns,
      displayData,
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