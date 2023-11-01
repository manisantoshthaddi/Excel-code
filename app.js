var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
var request = require('request');

var indexRouter = require('./routes/index');
var instructionsRouter = require('./routes/instructions');
var usersRouter = require('./routes/users');
var commandsRouter = require('./routes/commands');
var taskpaneRouter = require('./routes/taskpane');
var logoutRouter = require('./routes/logout');
var supportRouter = require('./routes/support');
var viewFavRouter = require('./routes/viewfav');
const { error } = require('console');

var app = express();
//app.use(cors())

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());

app.use('/images', express.static(path.join(__dirname, 'public/images')));
app.use('/javascripts', express.static(path.join(__dirname, 'public/javascripts')));
app.use('/stylesheets', express.static(path.join(__dirname, 'public/stylesheets')));

app.use('/', indexRouter);
app.use('/instructions', instructionsRouter);
app.use('/users', usersRouter);
app.use('/commands', commandsRouter);
app.use('/taskpane', taskpaneRouter);
app.use('/logout', logoutRouter);
app.use('/support', supportRouter);
app.use('/viewfav', viewFavRouter);


app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

app.post('/accesstoken', function(req,res){
  //let code = req.body.code;
  let api  = req.body;
  //res.send(api)
  //console.log('code ' + api.code);
  //var redirectUri = "https://excel-addin-prd.azurewebsites.net/commands";
  var args =  {
                  form:
                  {
                      'redirect_uri': api.ru,
                      'client_id': api.ci,
                      'client_secret': api.cs,
                      'code': `${api.code}`,
                      'grant_type':'authorization_code'
                  },
                  
                  rejectUnauthorized: false
              }

     // Build the URL to call for getting access token
      var urlapi = api.pu + api.ot
      //console.log('urlapi --> ' + urlapi);
      try{      
      request.post(urlapi, args, function(err,data,response){
              if(err)
              {
                  //console.log('error');
                  //console.log(err);
              }
              else{
                 //console.log(response);
                var outMessage = JSON.parse(response);
                res.send(outMessage);
              }
      });
    }
    catch(error){
      res.status(500)
      res.send(error)
    }
});

module.exports = app;
