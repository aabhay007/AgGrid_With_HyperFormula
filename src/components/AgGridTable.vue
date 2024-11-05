<template>
    <div class="mainapp">
        <div class="top-header-main">
            <h1>AG Grid with HyperFormula Integration</h1>
            <button @click="openModal">Add Custom Column</button>
        </div>
        <div class="theme_alpineLayout">
            <ag-grid-vue class="ag-theme-alpine" :columnDefs="columnDefs" :defaultColDef="defaultColDef"
                :rowData="rowData" :pagination="false" :getRowId="getRowNodeId" @grid-ready="onGridReady"
                @cell-value-changed="onCellValueChanged" />
        </div>
    </div>
    <!-- Modal -->
    <div v-if="isModalVisible" class="modal-overlay">
        <div class="modal">
            <h2>Add Custom Column</h2>
            <input type="text" v-model="newColumnName" placeholder="Enter column name" />
            <button @click="addCustomColumn">Add Column</button>
            <button @click="closeModal">Cancel</button>
        </div>
    </div>
</template>

<script lang="ts">
import { defineComponent, onMounted, ref } from 'vue';
import { AgGridVue } from 'ag-grid-vue3';
import type { ColDef, GridApi, GetRowIdFunc } from 'ag-grid-community';
import 'ag-grid-community/styles/ag-grid.css';
import 'ag-grid-community/styles/ag-theme-alpine.css';
import HyperFormula from 'hyperformula';
import Swal from 'sweetalert2';
import debounce from 'lodash/debounce';

interface RowData {
    id: number;
    [key: string]: any;
}

export default defineComponent({
    name: 'App',
    components: {
        AgGridVue
    },
    setup() {
        const columnDefs = ref<ColDef[]>([]);
        const rowData = ref<RowData[]>([]);
        const defaultColDef = {
            sortable: true,
            filter: false,
            resizable: true,
            editable: true,
        };
        const gridApi = ref<GridApi | null>(null);
        const loading = ref(false);
        const isModalVisible = ref(false);
        const newColumnName = ref('');
        const chunkSize = 1000;
        let currentChunk = 0;

        const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
        let sheetId: number | undefined = undefined;

        const getRowNodeId: GetRowIdFunc = (params) => params.data.id.toString();

        onMounted(() => {
            initializeSheet();
            generateData();
        });

        const initializeSheet = () => {
            try {
                const sheetName = 'Sheet1';
                hf.addSheet(sheetName);
                sheetId = hf.getSheetId(sheetName);
            } catch (error) {
                console.error('Error initializing sheet:', error);
            }
        };

        const generateData = () => {
            const columnDefinitionArray: ColDef[] = [];

            for (let i = 0; i < 600; i++) {
                columnDefinitionArray.push({
                    field: `col_${i + 1}`,
                    headerName: `Col ${i + 1}`,
                });
            }

            columnDefs.value = columnDefinitionArray;

            loadNextChunk();
        };

        const getRandomInt = (min: number, max: number): number => {
            return Math.floor(Math.random() * (max - min + 1)) + min;
        };

        const loadNextChunk = (): Promise<void> => {
            return new Promise((resolve) => {
                if (currentChunk * chunkSize >= 20000) {
                    resolve();
                    return;
                }

                loading.value = true;
                const start = currentChunk * chunkSize;
                const end = Math.min(start + chunkSize, 20000);
                const nextChunk: RowData[] = [];

                for (let i = start; i < end; i++) {
                    const row: RowData = { id: i };
                    row["col_1"] = "Apple";
                    row["col_2"] = i;
                    row["col_3"] = 0;

                    for (let j = 3; j < 600; j++) {
                        row[`col_${j + 1}`] = getRandomInt(1, 1000);
                    }

                    applyFormulaToRow(row, i);
                    nextChunk.push(row);
                }

                if (gridApi.value) {
                    gridApi.value.applyTransactionAsync({ add: nextChunk }, () => {
                        loading.value = false;
                        currentChunk++;
                        resolve();
                    });
                } else {
                    rowData.value = rowData.value.concat(nextChunk);
                    loading.value = false;
                    currentChunk++;
                    resolve();
                }
            });
        };

        const loadAllChunks = async () => {
            const totalChunks = Math.ceil(20000 / chunkSize);

            for (let i = 0; i < totalChunks; i++) {
                await loadNextChunk();
            }
        };

        const onGridReady = async (params: { api: GridApi }) => {
            gridApi.value = params.api;

            const observer = new IntersectionObserver((entries) => {
                if (entries[0].isIntersecting) {
                   // loadNextChunk();
                }
            }, { threshold: 0.5 });

            const gridElement = document.querySelector('.ag-theme-alpine');
            if (gridElement) {
                observer.observe(gridElement);
            }
            await loadAllChunks();
        };


        const openModal = () => {
            isModalVisible.value = true;
        };

        const closeModal = () => {
            isModalVisible.value = false;
            newColumnName.value = '';
        };

        const addCustomColumn = () => {
            if (newColumnName.value) {
                columnDefs.value.push({
                    field: newColumnName.value,
                    headerName: newColumnName.value,
                    editable: true,
                });
                Swal.fire({
                    title: `${newColumnName.value} Added Successfully!`,
                    icon: 'success',
                    confirmButtonText: 'OK'
                });
                closeModal();
            }
        };

        const applyFormulaToRow = (row: RowData, rowIndex: number) => {
            if (sheetId === undefined) return;

            const val1 = row['col_1'];
            const formula = `=IF(A${rowIndex + 1}="Apple", "NASDAQ", "FTSE")`;

            hf.setCellContents({ sheet: sheetId, col: 0, row: rowIndex }, [[val1]]);
            hf.setCellContents({ sheet: sheetId, col: 2, row: rowIndex }, [[formula]]);

            row['col_3'] = hf.getCellValue({ sheet: sheetId, col: 2, row: rowIndex }) || '';
        };

        const debouncedCellValueChanged = debounce((event: any) => {
            if (event.colDef.field === 'col_1') {
                const updatedRow = event.data;
                applyFormulaToRow(updatedRow, event.rowIndex);
                gridApi.value?.applyTransaction({ update: [updatedRow] });
            }
        }, 200);

        return {
            columnDefs,
            rowData,
            defaultColDef,
            gridApi,
            loading,
            isModalVisible,
            newColumnName,
            chunkSize,
            getRowNodeId,
            openModal,
            closeModal,
            addCustomColumn,
            onGridReady,
            onCellValueChanged: debouncedCellValueChanged,
        };
    },
});
</script>

<style>
#app {
    font-family: Avenir, Helvetica, Arial, sans-serif;
    -webkit-font-smoothing: antialiased;
    -moz-osx-font-smoothing: grayscale;
    text-align: center;
    color: #2c3e50;
    margin-top: 20px;
}

.modal-overlay {
    position: fixed;
    top: 0;
    left: 0;
    width: 100%;
    height: 100%;
    background-color: rgba(0, 0, 0, 0.5);
    display: flex;
    justify-content: center;
    align-items: center;
    z-index: 1000;
}

.modal {
    background-color: white;
    padding: 20px;
    border-radius: 10px;
    text-align: center;
    z-index: 1001;
}

.modal h2 {
    margin-bottom: 20px;
}

.modal input {
    margin-bottom: 10px;
    padding: 10px;
    width: 80%;
    border-radius: 5px;
    border: 1px solid #ccc;
}

.modal button {
    padding: 10px 20px;
    margin: 5px;
    border-radius: 5px;
    border: none;
    cursor: pointer;
}

.modal button:first-of-type {
    background-color: #4CAF50;
    color: white;
}

.modal button:last-of-type {
    background-color: #f44336;
    color: white;
}

.mainapp {
    width: 100%;
    padding: 0 15px 15px;
}

div#app {
    display: block;
    width: 100%;
}

.top-header-main {
    display: flex;
    margin-bottom: 6px;
    justify-content: space-between;
}

.top-header-main button {
    background-color: #333333;
    color: #fff;
    padding: 2px 10px;
    border-radius: 7px;
    line-height: inherit;
    cursor: pointer;
    border: 1px solid #333;
}

.top-header-main button:hover {
    background: #fff;
    color: #333;
}

.top-header-main h1 {
    font-size: 28px;
}
</style>