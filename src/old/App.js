import React, { Component } from 'react';
import { Helmet } from 'react-helmet';
import { ChangeCell } from './ChangeCell';

class App extends Component {
  constructor(props) {
    super(props);
    this.excelApiStatusID = setInterval(() => {
      if (!!window.fin.desktop.ExcelService) {
        this.clearIntervalAndStartService();
      }
    }, 100);
  }

  state = {
    currentBooks: [],
    currentSheets: [],
  };

  clearIntervalAndStartService = () => {
    console.log('clearIntervalAndStartService');
    clearInterval(this.excelApiStatusID);
    window.fin.desktop.ExcelService.init()
      .then(this.checkConnectionStatus)
      .then(() => console.log('Excel Service ready'))
      .catch((err) => console.error('Error: ', err));
  }

  checkConnectionStatus = () => {
    window.fin.desktop.Excel.getConnectionStatus((connected) => {
      connected
        ? console.log('Connected to Excel')
        : console.log('Excel not connected');
    })
  }

  connectToExcel = () => {
    console.log('connectToExcel');
    window.fin.desktop.Excel.run();
  }

  getBooks = () => {
    const bookPromise = window.fin.desktop.Excel.getWorkbooks();
    bookPromise.then((result) => {
      if (result.length) {
        this.setState({ currentBooks: result });
      } else {
        console.log('Could not find any open workbooks');
      }
    });
  }

  getSheets = () => {
    const { currentBooks } = this.state;
    console.log('this', this);
    if (currentBooks.length) {
      const sheetPromise = window.fin.desktop.Excel.getWorkbookByName(currentBooks[0].name).getWorksheets();
      sheetPromise.then((result) => {
        console.log('received: ', result);
        this.setState({ currentSheets: result });
      });
    }
  }

  render() {
    const { currentBooks, currentSheets } = this.state;
    return (
      <>
        <Helmet>
          <script src="https://openfin.github.io/excel-api-example/client/fin.desktop.Excel.js"></script>
        </Helmet>
        <button onClick={this.connectToExcel}>Click to connect</button>
        <button onClick={this.getBooks}>Click to get books</button>
        <button onClick={this.getSheets}>Click to get sheets</button>
        {currentBooks.length &&
          <ul>
            {currentBooks.map((book) => <li>{book.name}</li>)}
          </ul>
        }
        {currentSheets.length &&
          <>
            <ul>
              {currentSheets.map((book) => <li>{book.name}</li>)}
            </ul>
            <ChangeCell worksheet={currentSheets[0]} />
          </>
        }
      </>
    );
  }
}

export default App;
