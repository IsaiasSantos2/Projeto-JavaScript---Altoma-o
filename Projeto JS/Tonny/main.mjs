import { exec } from 'child_process';
import path from 'path';
import XLSX from 'xlsx';
import robot from 'robotjs';
import clipboardy from 'clipboardy';
process.stdin.setEncoding('utf-8');
process.stdout.setEncoding('utf-8');
 //const planilha = 'C:\\Users\\mundo\\Desktop\\produtos_ficticios.xlsx'
const planilha = 'C:\\Users\\mundo\\Desktop\\Meus Estudos\\Altomações JS\\produtos_ficticios.xlsx';

function preencherCampo(x, y, valor) {
    console.log('Valor a ser inserido:', valor);

    // Mover o mouse para as coordenadas do campo
    robot.moveMouse(x, y);

    // Clicar no campo para selecioná-lo
    robot.mouseClick();

    // Limpar o campo selecionado (se necessário)
    robot.keyTap('a', ['control']); // Selecionar tudo
    robot.keyTap('delete'); // Apagar

    // Escrever o valor diretamente no campo
    robot.typeString(valor);

    // Pressionar a tecla Tab para sair do campo
    robot.keyTap('tab');
}


function abrirPlanilha() {
    console.log('Planilha já deve estar aberta antes de executar o código.');
}

function aguarde(segundos) {
    return new Promise(resolve => {
        setTimeout(resolve, segundos * 1000); // Converte segundos em milissegundos
    });
}

async function iniciarProcesso() {

    const workbook = XLSX.readFile(planilha);
    const sheet_produtos = workbook.Sheets[workbook.SheetNames[0]];
    const dados = XLSX.utils.sheet_to_json(sheet_produtos);

    for (let index = 0; index < dados.length; index++) {
        const linha = dados[index];
        // Obtenha os valores de cada campo da linha
        const nomeProduto = linha['Nome do produto'];
        preencherCampo(1158, 348, nomeProduto);
        const descricao = linha['Descrição'];
        preencherCampo(1150, 437, descricao);
        const categoria = linha['Categoria'];
        preencherCampo(1190, 567, categoria);
        const codigodoproduto = linha['Código do produto'];
        preencherCampo(1167, 654, codigodoproduto);
        const peso = linha['Peso (em kg)'];
        const pesoString = peso.toString();
        preencherCampo(1149, 737, pesoString);
        const dimensoes = linha['Dimensões (L x A x P)'];
        preencherCampo(1166, 826, dimensoes);

        robot.moveMouse(1126, 878);
        //aqui tem um botão//
        robot.mouseClick();

        //aguarde 5 segundos
        await aguarde(3);

        //INICIE ETAPA 2
        const preco = linha['Preço'];

        const stringPreco = preco.toString();
        preencherCampo(1166, 390, stringPreco);

        const quantidadeemestoque = linha['Quantidade em estoque'];

        const quantidadeemestoquestring = quantidadeemestoque.toString()
        preencherCampo(1166, 475, quantidadeemestoquestring);

        const dataDeValidade = linha['Data de validade'];
        preencherCampo(1166, 560, dataDeValidade);

        const cor = linha['Cor'];
        preencherCampo(1166, 651, cor);

        const tamanho = linha['Tamanho'];
        robot.moveMouse(1134, 732);
        if (tamanho == 'Pequeno') {
            robot.mouseClick()
            robot.moveMouse(1134, 759)
            robot.mouseClick()
        } else if (tamanho == 'Médio') {
            robot.mouseClick()
            robot.moveMouse(1134, 780)
            robot.mouseClick()
        } else if (tamanho == 'Grande') {
            robot.mouseClick()
            robot.moveMouse(1134, 805)
            robot.mouseClick()
        }
        const material = linha['Material'];
        preencherCampo(1175, 818, material);
        robot.moveMouse(1119, 870)
        //botão aqui
        robot.mouseClick()
        await aguarde(3); // Exemplo de aguardar 2 segundos
        //ETAPA 3 - 
        const fabricante = linha['Fabricante'];
        preencherCampo(1165, 407, fabricante)
        const paisDeOrigem = linha['País de origem'];
        preencherCampo(1143, 497, paisDeOrigem)
        const observacoes = linha['Observações'];
        preencherCampo(1137, 593, observacoes)
        const codigoDeBarras = linha['Código de barras'];
        preencherCampo(1139, 716, codigoDeBarras)
        const localizacaoNoArmazem = linha['Localização no armazém'];
        preencherCampo(1138, 794, localizacaoNoArmazem)
        robot.moveMouse(1131,853)
        //BOTão aqui
        robot.mouseClick()
        await aguarde(2)
        robot.moveMouse(1612,184)
        //botao ok
        robot.mouseClick()
        await aguarde(2)
        //botao ok
        robot.mouseClick()
        robot.moveMouse(1431,615)
        //botão adicionar mais um

        if(index < dados.length){
            robot.mouseClick()
        }
        await aguarde(3)
        console.log("DEU BOM!")
    }
}

iniciarProcesso().then(() => {
    console.log('Processo concluído.');
})

.catch(error => {
    console.error('Ocorreu um erro:', error);
});
