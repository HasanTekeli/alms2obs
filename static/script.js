const fs = require("fs");
const pyshell = require("python-shell")

document.querySelector('form')
    .addEventListener('submit', function(event) {
        event.preventDefault();
        const folder_path = document.querySelector('input').files[0].path.replace(/\/[^\/]+$/, '');
        //console.log(folder_path)
        document.getElementById("ulist").innerHTML = folder_path;
        const readDir = fs.readdirSync(folder_path);
        let text = "";
        for(let filePath of readDir) {
            //console.log(filePath);
            text += filePath + "<br>";
        }
        document.getElementById("olist").innerHTML = text;
    })

let pyshell_options = {
    pythonOptions: ['-u'],
    args: [folder_path]
}
function modify_excel() {
    pyshell.run('deneme.py', pyshell_options)

}

document.addEventListener('DOMContentLoaded', function() {
    
    var url = 'http://127.0.0.1:5001/GUI-is-still-open'; 
    fetch(url, { mode: 'no-cors'});
    setInterval(function(){ fetch(url, { mode: 'no-cors'});}, 5000)();

});