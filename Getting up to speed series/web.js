const http = require('http');

const server = http.createServer((req, res) => {
    res.end("Hello, I'm now running in your browser.");
});

server.listen(8080, (error) => {
  if (error) {
      return console.log('ERROR: ', error)
  }

  console.log('Server is listening on 8080. Navigate to: http://localhost:8080')
});