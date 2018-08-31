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
    Author: 'Tuleap',
    CreatedDate: new Date()
  };

  wb.SheetNames.push('srs');

  var artifactPreviousLength = 0;
  var longestArtifactIndex = 0;
  var idx = 0;

  // Check which is the artifact with more fields and retrive its index
  artifacts.forEach((artifact) => {
    if (artifact.values.length > artifactPreviousLength) {
      longestArtifactIndex = idx;
      artifactPreviousLength = artifact.values.length;
    }
    idx += 1;
  });

  var headers = [];
  var fieldsIds = [];
  artifacts[longestArtifactIndex].values.forEach((field) => {
    if (field.label !== undefined) {
      headers.push(field.label);
      fieldsIds.push(field.field_id);
    } else {
      headers.push('NEF'); // NEF = non-existing field
    }

  });

  var wsData = [];

  wsData.push(headers);

  artifacts.forEach((data) => {
    var row = [];
    var field = null;
    var fields = data.values;

    for (var i = 0; i < fieldsIds.length; i ++) {

      field = _.filter(fields, o => o.field_id == fieldsIds[i])[0];

      if (field !== undefined) {

        if (field.value !== undefined) {
          if (Array.isArray(field.value)) {
            var data = [];

            field.value.forEach((obj) => {
              if (obj.ref !== undefined) {
                data.push(obj.ref);
              } else {
                data.push('NN');
              }
            });

            row.push(_.join(data, ','));
          } else if (typeof field.value === 'object') {
            if (field.value.real_name !== undefined) {
              row.push(field.value.real_name);
            } else if (field.value.url !== undefined) {
              row.push(field.value.url)
            }
          } else {
            row.push(field.value);
          }
        } else if (field.values !== undefined) {
          var data = [];

          field.values.forEach((obj) => {
            if (obj.label !== undefined) {
              data.push(obj.label);
            } else if (obj.real_name !== undefined) {
              data.push(obj.real_name);
            } else {
              data.push('N/A');
            }
          });

          row.push(_.join(data, ','));
        } else if (field.links !== undefined) {
          var links = [];

          field.links.forEach((obj) => {
            links.push(obj.uri);
          });

          row.push(_.join(links, ','));
        } else {
          row.push('N/A');
        }
      } else {
        row.push('N/A');
      }
    }

    wsData.push(row);
  });

  var ws = XLSX.utils.aoa_to_sheet(wsData);
  wb.Sheets['srs'] = ws;

  let dt = moment().format('YYYYMMDD-HHmm');
  XLSX.writeFile(wb, `tuleap-artifacts-${dt}.xlsx`);

}).catch((err) => {
  console.log(err);
});
