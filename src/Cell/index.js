import React, { useState } from 'react';
import classNames from 'classnames';
import styles from './styles.module.scss';

export const Cell = ({ storeValue, initValue = '' }) => {
  const [value, setValue] = useState(initValue);
  const [focus, setFocus] = useState(false);

  const inputClass = classNames(
    styles.cell,
    {
      [styles.hasFocus]: focus,
    }
  );
  const handleChange = (e) => {
    const val = e.target.value;
    setValue(val);
    storeValue(val);
  }
  return (
    <input
      className={inputClass}
      type="text"
      value={value}
      onBlur={() => setFocus(false)}
      onFocus={() => setFocus(true)}
      onChange={handleChange}
    />
  );
}
