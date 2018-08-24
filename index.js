const XLSX = require('xlsx');
const https = require('https');
const axios = require('axios');
const _ = require('lodash');
const moment = require('moment');

const baseUrl = 'https://tuleap.siscog/api/';

const instance = axios.create({
  httpsAgent: new https.Agent({
    rejectUnauthorized: false
  }),
  auth: {
    username: 'admin',
    password: 'Tq26o49wNlXWaEp'
  }
});

instance.get(baseUrl + 'trackers/17/artifacts?values=all').then((res) => {
  let artifacts = _.filter(res.data, 'values');

  var wb = XLSX.utils.book_new();

  wb.Props = {
    Title: 'SRS',
    Subject: 'Export file',
    Author: 'Hugo S. Dias',
    CreatedDate: new Date()
  };

  wb.SheetNames.push('srs');

  var wsData = [];

  artifacts.forEach((data) => {
    var rowIdx = 1;
    var newRow = [];
    data.values.forEach((col) => {
      var colIdx = 0;

      if (col.value !== undefined) {
        newRow.push(col.value);
      } else if (col.values !== undefined) {
        var labels = [];

        col.values.forEach((obj) => {
          labels.push(obj.label);
        });
        
        newRow.push(_.join(labels, ','));
      }

      colIdx ++;
    });

    console.log(newRow);
    wsData.push(newRow);

    rowIdx ++;
  });

  var ws = XLSX.utils.aoa_to_sheet(wsData);
  wb.Sheets['srs'] = ws;

  let dt = moment().format('YYYYMMDD-HHmm');
  XLSX.writeFile(wb, `tuleap-artifacts-${dt}.xlsx`);
  // res.data[0].values.forEach((val) => {
  //   if (val.value !== undefined) {
  //     console.log(val.label, '=', val.value);
  //   } else if (val.values !== undefined) {
  //     console.log(val.label, '=', val.values);
  //   } else if (val.links !== undefined) {
  //     console.log(val.label, '=', val.links)
  //   }
  // });
}).catch((err) => {
  console.log(err);
});
