const express = require('express');
const path = require('path');
const { getDefaultData, availableIcons } = require('./default-data');

const app = express();
const PORT = process.env.PORT || 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

function renderPage(req, res) {
  const data = getDefaultData();
  res.render('index', { data, availableIcons });
}

app.get('/', renderPage);
app.get('/digital-transaction-snapshot', renderPage);

app.get('/api/default-data', (req, res) => {
  res.json(getDefaultData());
});

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`Digital Transaction Value Snapshot: http://localhost:${PORT}/digital-transaction-snapshot`);
});

module.exports = app;
