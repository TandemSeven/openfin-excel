import React, { Component } from 'react';
import { Helmet } from 'react-helmet';
import { Cell } from './Cell';
import styles from './styles.module.scss';

const COLUMN_KEYS = ['A', 'B', 'C', 'D', 'E'];
const ROW_KEYS = ['1', '2', '3', '4', '5'];
const INITIAL_CELL_GRID = [
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
  ['', '', '', '', ''],
];

export default class App extends Component {
  constructor(props) {
    super(props);
    this.excelApiStatusID = setInterval(() => {
      if (!!window.fin.desktop.ExcelService) {
        this.clearIntervalAndStartService();
      }
    }, 100);
  }

  state = {
    sheet: null,
    cellGrid: INITIAL_CELL_GRID,
  }

  // Clear the intervalId
  clearIntervalAndStartService = () => {
    console.log('clearIntervalAndStartService');
    clearInterval(this.excelApiStatusID);
    window.fin.desktop.ExcelService.init()
      .then(this.checkConnectionStatus)
      .then(() => console.log('Excel Service ready'))
      .then(this.connectToExcel)
      .then(this.createBook)
      .then(this.getSheet)
      .then(this.activateSheet)
      .catch((err) => console.error('Error: ', err));
  }

  // Check the status off the Excel service connection
  checkConnectionStatus = () => {
    window.fin.desktop.Excel.getConnectionStatus((connected) => {
      connected
        ? console.log('Connected to Excel')
        : console.log('Excel not connected');
    })
  }

  // Boot up the Excel program
  connectToExcel = () => {
    console.log('connectToExcel');
    return window.fin.desktop.Excel.run();
  }

  createBook = () => {
    console.log('createBook');
    return window.fin.desktop.Excel.addWorkbook();
  }

  getSheet = (workbook) => {
    console.log('getSheet from', workbook);
    return workbook.getWorksheets();
  }

  activateSheet = (sheet) => {
    console.log('activateSheet', sheet[0]);
    this.setState({ sheet: sheet[0] });
    return sheet[0].activate();
  }

  // Set latest value in state
  storeValue = (x, y) => (val) => {
    this.setState(({ cellGrid }) => {
      const newGrid = [...cellGrid];
      newGrid[x][y] = val;
      return newGrid;
    });
  }

  // example using async to update state with result of promise
  getBooks = async () => {
    this.setState({ books: await window.fin.desktop.Excel.getWorkbooks() });
  }

  pushToSheet = async () => {
    const { cellGrid, sheet } = this.state;
    console.log('Saving: ', cellGrid);
    console.log('to:', sheet);
    const result = await sheet.setCells(cellGrid, 'A1');
    console.log('Response: ', result);
  }

  render() {
    return (
      <>
        <Helmet>
          <script src="https://openfin.github.io/excel-api-example/client/fin.desktop.Excel.js"></script>
        </Helmet>
        <main className={styles.container}>
          <div className={styles.cellGrid}>
            {
              COLUMN_KEYS.map((colVal, colIdx) => (
                ROW_KEYS.map((rowVal, rowIdx) => (
                  <Cell
                    cellKey={`${colVal}${rowVal}`}
                    storeValue={this.storeValue(colIdx, rowIdx)}
                    key={`${colIdx}-${rowIdx}`}
                  />
                ))
              ))
            }
          </div>
          <button type="button" onClick={() => this.pushToSheet()}>Transfer</button>
        </main>
      </>
    )
  }
};
