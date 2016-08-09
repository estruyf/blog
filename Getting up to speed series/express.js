const express = require('express')
var app = express()
 
app.get('/', function (req, res) {
  res.send("<strong>Hello, I'm now running in your browser with the express package.</strong>")
})
 
app.listen(8080, (error) => {
  if (error) {
      return console.log('ERROR: ', error)
  }

  console.log('Server is listening on 8080. Navigate to: http://localhost:8080')
});