<template>
  <div class="excel-table">
    <div class="file-upload">
      <input type="file" @change="handleFileUpload" accept=".xlsx,.csv" />
      <button @click="uploadFile" :disabled="!file">Upload</button>
    </div>

    <div v-if="error" class="error-message">
      {{ error }}
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

export default defineComponent({
  name: 'ExcelTable',
  setup() {
    const columns = ref([])
    const tableData = reactive([])
    const parser = new Parser()
    const file = ref(null)
    const error = ref(null)

    const handleFileUpload = (event) => {
      file.value = event.target.files[0]
      error.value = null  // Clear any previous errors
    }

    const uploadFile = async () => {
      if (!file.value) return
      
      const formData = new FormData()
      formData.append('file', file.value)

      try {
        const response = await axios.post('http://localhost:8000/upload-excel/', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        })
        console.log('File uploaded successfully:', response.data)
        
        // Update tableData and columns with the response data
        if (response.data && response.data.data) {
          tableData.splice(0, tableData.length, ...response.data.data)
          columns.value = response.data.column_names
        }
      } catch (error) {
        console.error('Error uploading file:', error)
        error.value = "Error uploading file. Please try again."
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
      uploadFile,
      file,
      error,
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
.file-upload input[type="file"] {
  margin-right: 1rem;
}
.error-message {
  color: red;
  margin-bottom: 1rem;
}
</style>