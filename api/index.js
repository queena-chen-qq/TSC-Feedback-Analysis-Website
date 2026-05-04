const express = require('express');
const { DynamoDBClient } = require('@aws-sdk/client-dynamodb');
const { DynamoDBDocumentClient, PutCommand, QueryCommand, ScanCommand, DeleteCommand } = require('@aws-sdk/lib-dynamodb');

const app = express();
app.use(express.json({ limit: '10mb' }));

const client = new DynamoDBClient({
  region: process.env.AWS_REGION || 'us-east-1',
  credentials: {
    accessKeyId: process.env.MY_AWS_ACCESS_KEY_ID,
    secretAccessKey: process.env.MY_AWS_SECRET_ACCESS_KEY
  }
});
const ddb = DynamoDBDocumentClient.from(client);

const FEEDBACK_TABLE = 'tsc-feedbacks';
const PEER_TABLE = 'tsc-peer-reviews';

// --- Feedback APIs ---

// Save feedbacks (batch from frontend)
app.post('/api/feedbacks', async (req, res) => {
  try {
    const { batch, records } = req.body;
    if (!batch || !records?.length) return res.status(400).json({ error: '無效資料' });

    for (const record of records) {
      await ddb.send(new PutCommand({
        TableName: FEEDBACK_TABLE,
        Item: { pk: batch, sk: String(record.id), ...record }
      }));
    }
    res.json({ message: `成功儲存 ${records.length} 筆`, count: records.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Get all batches
app.get('/api/feedbacks/batches', async (req, res) => {
  try {
    const result = await ddb.send(new ScanCommand({
      TableName: FEEDBACK_TABLE,
      ProjectionExpression: 'pk'
    }));
    const batches = [...new Set((result.Items || []).map(i => i.pk))].sort();
    res.json(batches);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// Get feedbacks by batch
app.get('/api/feedbacks', async (req, res) => {
  try {
    const { batch } = req.query;
    if (!batch) {
      const result = await ddb.send(new ScanCommand({ TableName: FEEDBACK_TABLE }));
      return res.json(result.Items || []);
    }
    const result = await ddb.send(new QueryCommand({
      TableName: FEEDBACK_TABLE,
      KeyConditionExpression: 'pk = :b',
      ExpressionAttributeValues: { ':b': batch }
    }));
    res.json(result.Items || []);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// Delete all feedbacks
app.delete('/api/feedbacks', async (req, res) => {
  try {
    const result = await ddb.send(new ScanCommand({
      TableName: FEEDBACK_TABLE, ProjectionExpression: 'pk, sk'
    }));
    for (const item of (result.Items || [])) {
      await ddb.send(new DeleteCommand({ TableName: FEEDBACK_TABLE, Key: { pk: item.pk, sk: item.sk } }));
    }
    res.json({ message: '已清除' });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- Peer Review APIs ---

app.post('/api/peers', async (req, res) => {
  try {
    const { batch, records } = req.body;
    if (!batch || !records?.length) return res.status(400).json({ error: '無效資料' });

    for (const record of records) {
      await ddb.send(new PutCommand({
        TableName: PEER_TABLE,
        Item: { pk: batch, sk: String(record.id), ...record }
      }));
    }
    res.json({ message: `成功儲存 ${records.length} 筆`, count: records.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/peers/batches', async (req, res) => {
  try {
    const result = await ddb.send(new ScanCommand({
      TableName: PEER_TABLE,
      ProjectionExpression: 'pk'
    }));
    const batches = [...new Set((result.Items || []).map(i => i.pk))].sort();
    res.json(batches);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/peers', async (req, res) => {
  try {
    const { batch } = req.query;
    if (!batch) {
      const result = await ddb.send(new ScanCommand({ TableName: PEER_TABLE }));
      return res.json(result.Items || []);
    }
    const result = await ddb.send(new QueryCommand({
      TableName: PEER_TABLE,
      KeyConditionExpression: 'pk = :b',
      ExpressionAttributeValues: { ':b': batch }
    }));
    res.json(result.Items || []);
  } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete('/api/peers', async (req, res) => {
  try {
    const result = await ddb.send(new ScanCommand({
      TableName: PEER_TABLE, ProjectionExpression: 'pk, sk'
    }));
    for (const item of (result.Items || [])) {
      await ddb.send(new DeleteCommand({ TableName: PEER_TABLE, Key: { pk: item.pk, sk: item.sk } }));
    }
    res.json({ message: '已清除' });
  } catch (err) { res.status(500).json({ error: err.message }); }
});

module.exports = app;
