var createError = require('http-errors');
var express = require('express');
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
const ejs = require('ejs');
var indexRouter = require('./routes/index');
var usersRouter = require('./routes/users');
const cors = require('cors');

var app = express();
var todoRouter=require('./routes/users')  
app.use('/test',todoRouter)  
// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(logger('dev'));
app.use(express.json());
app.use(express.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));
// Assuming you have an 'app' instance of Express
// app.use('/public', express.static(path.join(__dirname, 'public')));

// Add this to your Express app to serve static files
app.use('/images', express.static(path.join(__dirname, 'public/images')));

app.use('/', indexRouter);
app.use('/users', usersRouter);

app.use(cors());


var salaryRouter = require('./routes/salary');



app.use('/salary', salaryRouter);




// catch 404 and forward to error handler
app.use(function(req, res, next) {
  next(createError(404));
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
