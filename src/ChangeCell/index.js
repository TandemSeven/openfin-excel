import React, { useState } from 'react';

export const ChangeCell = ({ worksheet = {} }) => {
  const listenerWithArgs = (e) => console.log('listenerWithArgs', e.data.range);
  worksheet.addEventListener('selectionChanged', listenerWithArgs);
  const [newValue, setNewValue] = useState('');
  const [result, setResult] = useState(false);
  // This block is for reading data from a cell
  // const [cellData, setCellData] = useState(null);
  // const getEnteredCell = (enteredCell) => {
  //   setCell(enteredCell);
  //   worksheet
  //     && worksheet.getCells
  //     && worksheet.getCells(enteredCell, 10, 10) // cell, row offset, column offset
  //       .then((result) => {
  //         if (result.length) {
  //           setCellData(result[0][0].value)
  //         }
  //       })
  //       .catch((err) => console.log('error: ', err));
  // };
  const setCellValue = () => {
    console.log('sending', newValue);
    worksheet
      && worksheet.setCells
      && worksheet.setCells([[newValue]], 'A1')
        .then((response) => {
          console.log('response: ', response);
          setResult(true);
        })
        .catch((err) => console.log('error: ', err));
  }
  return (
    <>
      <p>Insert into cell A1:</p>
      <input type="text" onChange={(e) => setNewValue(e.target.value)} />
      <button type="button" onClick={() => setCellValue(newValue)}>Set value</button>
      {result && <p>Success!</p>}
      {/* BELOW code will read a value from a specific cell
      <input onChange={(e) => setInput(e.target.value)} value={input} />
      <button type="button" onClick={() => getEnteredCell(input)}>Click</button>
      {cell && <p>Current selection: {cell}</p>}
      {cellData && <p>Value: {cellData}</p>} */}
    </>
  );
}