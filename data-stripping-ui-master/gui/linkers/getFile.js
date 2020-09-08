let { PythonShell } = require('python-shell')
let {shell} = require('electron')
var path = require("path")

function submitFile() {
  const ALERT_SUCCESS_CLASS = 'alert alert-success';
  const ALERT_DANGER_CLASS = 'alert alert-danger';
  const SUCCESS_ALERT_MESSAGE = 'The file was successfully processed';
  const ERROR_ALERT_MESSAGE = 'There was a problem processing the file';

  document.getElementById('stripFileResponse').hidden = true;
  document.getElementById('spinner').hidden = false;

  function handleAlert(className, message) {
    document.getElementById('stripFileResponse').className = className;
    document.getElementById('stripFileResponse').innerHTML = message;
    document.getElementById('stripFileResponse').hidden = false;
  }

  var file = document.getElementById("myfile").files[0]
  
  if (file == null){
    handleAlert(ALERT_DANGER_CLASS, "Please, select a file")
    document.getElementById('spinner').hidden = true;
  }

  var options = {
    scriptPath: path.join(__dirname, '/../engine/'),
    args: [file.path, file.name]
  }

  let pyshell = new PythonShell('Stripping.py', options);

  pyshell.end(function (err, code, signal) {
    if (err) {
      handleAlert(ALERT_DANGER_CLASS, ERROR_ALERT_MESSAGE);
    }
    else {
      handleAlert(ALERT_SUCCESS_CLASS, SUCCESS_ALERT_MESSAGE);
    }
    document.getElementById('spinner').hidden = true;
  });
}

function openOutputFolder(){

  // shell.openItem('C:\Users\alexg\Data Stripping\engine\Results')
  shell.showItemInFolder('C:/Users/alexg/Data Stripping/engine/Results/.')

}