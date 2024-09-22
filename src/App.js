import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

function App() {
  const [data, setData] = useState([]);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = (e) => {
      const binaryStr = e.target.result;
      const workbook = XLSX.read(binaryStr, { type: 'binary' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      setData(jsonData);
      console.log(jsonData);
    };
    reader.readAsBinaryString(file);
  };

  const parseDate = (dateValue) => {
    if (typeof dateValue === 'string') {
      const [day, month, year] = dateValue.split('-');
      return `${year}${month}${day}`;
    } else if (dateValue instanceof Date) {
      const year = dateValue.getFullYear();
      const month = (`0${dateValue.getMonth() + 1}`).slice(-2);
      const day = (`0${dateValue.getDate()}`).slice(-2);
      return `${year}${month}${day}`;
    } else if (typeof dateValue === 'number') {
      const parsedDate = new Date((dateValue - 25569) * 86400 * 1000);
      const year = parsedDate.getFullYear();
      const month = (`0${parsedDate.getMonth() + 1}`).slice(-2);
      const day = (`0${parsedDate.getDate()}`).slice(-2);
      return `${year}${month}${day}`;
    } else {
      throw new Error(`Unexpected date format: ${dateValue}`);
    }
  };

  const generateTallyXML = () => {
    if (!data.length) return;

    const createXml = (data) => {
      const xmlDoc = document.implementation.createDocument('', '', null);
      const envelope = xmlDoc.createElement('ENVELOPE');

      const header = xmlDoc.createElement('HEADER');
      const tallyRequest = xmlDoc.createElement('TALLYREQUEST');
      tallyRequest.textContent = 'Import Data';
      header.appendChild(tallyRequest);
      envelope.appendChild(header);

      const body = xmlDoc.createElement('BODY');
      const importData = xmlDoc.createElement('IMPORTDATA');
      const requestDesc = xmlDoc.createElement('REQUESTDESC');
      const reportName = xmlDoc.createElement('REPORTNAME');
      reportName.textContent = 'Vouchers';
      requestDesc.appendChild(reportName);

      const staticVars = xmlDoc.createElement('STATICVARIABLES');
      const svcCompany = xmlDoc.createElement('SVCURRENTCOMPANY');
      svcCompany.textContent = 'Your Company Name';
      staticVars.appendChild(svcCompany);
      requestDesc.appendChild(staticVars);
      importData.appendChild(requestDesc);

      const requestData = xmlDoc.createElement('REQUESTDATA');

      data.forEach((row, index) => {
        const tallyMessage = xmlDoc.createElement('TALLYMESSAGE');
        const voucherType = row['Withdrawals'] ? 'Payment' : 'Receipt';
        const voucher = xmlDoc.createElement('VOUCHER');
        voucher.setAttribute('VCHTYPE', voucherType);
        voucher.setAttribute('ACTION', 'Create');
        voucher.setAttribute('OBJVIEW', 'Accounting Voucher View');

        const dateStr = parseDate(row['Date']);
        const date = xmlDoc.createElement('DATE');
        date.textContent = dateStr;
        voucher.appendChild(date);

        const vchType = xmlDoc.createElement('VOUCHERTYPENAME');
        vchType.textContent = voucherType;
        voucher.appendChild(vchType);

        const vchNumber = xmlDoc.createElement('VOUCHERNUMBER');
        vchNumber.textContent = (index + 1).toString();
        voucher.appendChild(vchNumber);

        const partyLedgerName = xmlDoc.createElement('PARTYLEDGERNAME');
        partyLedgerName.textContent = row['Ledger Name'];
        voucher.appendChild(partyLedgerName);

        const narration = xmlDoc.createElement('NARRATION');
        narration.textContent = row['Particulars'];
        voucher.appendChild(narration);

        const ledgerEntry1 = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
        const ledgerName1 = xmlDoc.createElement('LEDGERNAME');
        ledgerName1.textContent = row['Ledger Name'];
        ledgerEntry1.appendChild(ledgerName1);

        const isDeemedPositive1 = xmlDoc.createElement('ISDEEMEDPOSITIVE');
        isDeemedPositive1.textContent = voucherType === 'Receipt' ? 'No' : 'Yes';
        ledgerEntry1.appendChild(isDeemedPositive1);

        const amount1 = xmlDoc.createElement('AMOUNT');
        amount1.textContent = voucherType === 'Receipt' ? row['Deposits'] : `-${row['Withdrawals']}`;
        ledgerEntry1.appendChild(amount1);
        voucher.appendChild(ledgerEntry1);

        const ledgerEntry2 = xmlDoc.createElement('ALLLEDGERENTRIES.LIST');
        const ledgerName2 = xmlDoc.createElement('LEDGERNAME');
        ledgerName2.textContent = row['Bank Name'];
        ledgerEntry2.appendChild(ledgerName2);

        const isDeemedPositive2 = xmlDoc.createElement('ISDEEMEDPOSITIVE');
        isDeemedPositive2.textContent = voucherType === 'Receipt' ? 'Yes' : 'No';
        ledgerEntry2.appendChild(isDeemedPositive2);

        const amount2 = xmlDoc.createElement('AMOUNT');
        amount2.textContent = voucherType === 'Receipt' ? `-${row['Deposits']}` : row['Withdrawals'];
        ledgerEntry2.appendChild(amount2);
        voucher.appendChild(ledgerEntry2);

        tallyMessage.appendChild(voucher);
        requestData.appendChild(tallyMessage);
      });

      importData.appendChild(requestData);
      body.appendChild(importData);
      envelope.appendChild(body);
      xmlDoc.appendChild(envelope);

      const serializer = new XMLSerializer();
      return serializer.serializeToString(xmlDoc);
    };

    const xmlContent = createXml(data);
    const blob = new Blob([xmlContent], { type: 'application/xml' });
    saveAs(blob, 'TallyData.xml');
  };

  return (
    <div className="App">
      <h1>Excel to Tally XML Converter - Bank Statement</h1>
      <input type="file" onChange={handleFileUpload} />
      <button onClick={generateTallyXML}>Generate Tally XML</button>
    </div>
  );
}

export default App
