var express = require('express');
var router = express.Router();


router.get('/', function (req, res) {
    res.render('commands');
});

// router.get('/oauth/redirect', function (req, res) {
//     let code = req.query.code;
//     console.log("Code is " + code);
//     res.redirect('/commands');
// });

module.exports = router;
