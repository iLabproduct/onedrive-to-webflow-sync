require('dotenv').config();
const express = require('express');
const axios = require('axios');
const msal = require('@azure/msal-node');

const app = express();
const port = process.env.PORT || 3000;

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  }
};

const cca = new msal.ConfidentialClientApplication(msalConfig);
const tokenRequest = {
  scopes: ['https://graph.microsoft.com/.default'],
};

async function getAccessToken() {
  const response = await cca.acquireTokenByClientCredential(tokenRequest);
  return response.accessToken;
}

async function getExcelData(accessToken) {
  const response = await axios.get(`https://graph.microsoft.com/v1.0/me/drive/root:${process.env.FILE_PATH}:/workbook/worksheets/${process.env.WORKSHEET_NAME}/usedRange`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  return response.data.values;
}

function transformExcelToWebflowFormat(excelData) {
  const headers = excelData[0];
  return excelData.slice(1).map(row => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = row[index];
    });
    return item;
  });
}

async function updateWebflowCMS(item) {
  const webflowResponse = await axios.post(`https://api.webflow.com/collections/${process.env.COLLECTION_ID}/items`, {
    fields: {
      name: item.name,
      slug: item.name.toLowerCase().replace(/\s+/g, '-'),
      _archived: false,
      _draft: false,
      field1: item.field1,
      field2: item.field2,
    }
  }, {
    headers: {
      Authorization: `Bearer ${process.env.WEBFLOW_API_KEY}`,
      'Content-Type': 'application/json',
      'accept-version': '1.0.0'
    }
  });
  return webflowResponse.data;
}

app.get('/sync', async (req, res) => {
  try {
    const accessToken = await getAccessToken();
    const excelData = await getExcelData(accessToken);
    const transformedData = transformExcelToWebflowFormat(excelData);

    for (const item of transformedData) {
      await updateWebflowCMS(item);
    }

    res.status(200).send('Data synced to Webflow successfully');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error syncing data');
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
