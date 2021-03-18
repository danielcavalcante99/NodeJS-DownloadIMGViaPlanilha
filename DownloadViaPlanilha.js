//Autor: Daniel Henrique Cavalcante da Silva

/*Programinha para ler a coluna da planilha que tenha
URLS de imagens e fazer o download em massa*/

//npm install xlsx --save
const xlsx = require("xlsx");
//npm install https --save
const https = require('https');
const fs = require('fs');

function downloadViaPlanilha(Caminho, AbaPlanilha, ColunaLink, ColunaID) {

    var planilha = xlsx.readFile(Caminho, {cellDates:true});

    var aba = planilha.Sheets[AbaPlanilha];

    var dados = xlsx.utils.sheet_to_json(aba);

        dados.forEach(coluna => {


        var pasta = process.env.userprofile + '\\Downloads\\IMG - DownloadViaPlanilha';
        //Criando a pasta IMG - DownloadViaPlanilha na pasta Downloads
        if(!fs.existsSync(pasta)){
        fs.mkdirSync(pasta);
        }
    
        var arquivo = coluna[ColunaID];
        var req = https.get(coluna[ColunaLink], function(res){
        var fileStream = fs.createWriteStream(pasta+'\\'+arquivo+'.jpg');
        res.pipe(fileStream);
    
        fileStream.on("error", function(err){
            console.log("Error writting to the stream.");
            console.log(err);
        });
    
        fileStream.on("finish", function() {
            fileStream.close();
            console.log("Download: " + coluna[ColunaID]);

        req.on("error", function(err){
            console.log("Error downloading the file.");
            console.log(err);
        });
      }); 
    }); 
  });
}

//Passe as informações para os parâmetros dentro do parêntese da função abaixo
downloadViaPlanilha();