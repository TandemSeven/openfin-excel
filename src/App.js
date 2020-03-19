import React, { Component } from 'react';
import { Helmet } from 'react-helmet';
import { Cell } from './Cell';
import styles from './styles.module.scss';

const COLUMN_KEYS = ['A', 'B', 'C', 'D', 'E'];
const ROW_KEYS = ['1', '2', '3', '4', '5'];

// This is meant to be an in-memory representation of the entire grid of cells
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
  clearIntervalAndStartService = async () => {
    console.log('clearIntervalAndStartService');
    clearInterval(this.excelApiStatusID);
    try {
      // Start the service
      await window.fin.desktop.ExcelService.init();
      // Connect to Excel application
      await this.connectToExcel();

      // Confirm status of connection to Excel
      const connected = await this.checkConnectionStatus();
      if (connected) console.log('Connected to Excel');
      if (!connected) throw Error('Failed to establish connection');

      // Create a new book
      const book = await this.createBook();
      // Get the newly created sheets
      const sheets = await this.getSheet(book);
      // Activate the sheet to begin editing
      await this.activateSheet(sheets);
    } catch (error) {
      console.log('error: ', error);
    }
  }

  checkConnectionStatus = () => {
    return window.fin.desktop.Excel.getConnectionStatus();
  }

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
    this.setState({ book: workbook });
    return workbook.getWorksheets();
  }

  activateSheet = (sheets) => {
    console.log('activateSheet', sheets[0]);
    this.setState({ sheet: sheets[0] });
    return sheets[0].activate();
  }

  // Set latest value in state
  storeValue = (x, y) => (val) => {
    this.setState(({ cellGrid }) => {
      const newGrid = [...cellGrid];
      newGrid[x][y] = val;
      return newGrid;
    });
  }

  // Push the entire `cellGrid` to the `sheet` offset at cell `A1`
  pushToSheet = async () => {
    const { book, cellGrid, sheet } = this.state;
    console.log('Saving: ', cellGrid);
    console.log('to:', sheet);
    const result = await sheet.setCells(cellGrid, 'A1');
    console.log('Response: ', result);
    // Force save the book
    this.saveSheet(book)
  }

  saveSheet = async (book) => {
    await book.save();
  }

  render() {
    const { cellGrid } = this.state;
    return (
      <>
        {/* This script is required to use the Excel plugin */}
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
                    initValue={cellGrid[colIdx][rowIdx]}
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
