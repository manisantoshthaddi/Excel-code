var express = require('express');
var router = express.Router();

/* GET home page. */
router.get('/', function(req, res) {
  res.render('webauthtokenapi.js');
});

module.exports = router;
