fs = require('fs');
var runner = require("child_process");
  const q = require('q');

module.exports.render = render;


function render(WorkBook){
  
const temp1 =  __dirname + '/temp/'+Math.random();
const temp =  __dirname + '/temp/';
const filePath = temp+ WorkBook.name;
WorkBook.path  = filePath;

var jsonFile = temp + Math.random() + '.json';

var deferred = q.defer();

 fs.writeFile(jsonFile, JSON.stringify(WorkBook), function (err) {

  if (err) {
    deferred.reject(new Error(err));
  }
  
  //sudo apt-get install php5-cli
  
  runner.exec("php  convertor.php " + jsonFile,
      function (err, stdout, stderr) {
        console.log(err);
        console.log(stderr);
        console.log(stdout);
      if(!err){
        try {
          var buffer = fs.readFileSync(filePath);
          fs.unlinkSync(filePath);
          deferred.resolve({
            headers : {
              'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
              'Content-Disposition': 'attachment; filename='+ WorkBook.name ,
              'Content-Length': buffer.length
            },
            report : buffer
          });
        } catch (err) {
          deferred.reject(new Error(err));
        }
      }
      });
});

return deferred.promise;
}