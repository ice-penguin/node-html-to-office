var client = require("../index");

client.htmlStrToDocx('<p>html字符串</p>', {headerTitle:"标题",outputFile:__dirname + "/../file/test2"})
.then(res => {
    console.log(res)
})