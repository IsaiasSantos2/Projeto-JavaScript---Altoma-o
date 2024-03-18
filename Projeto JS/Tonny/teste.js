import { exec } from 'child_process';
import path from 'path';
import XLSX from 'xlsx';
import robot from 'robotjs';
import clipboardy from 'clipboardy';
const planilha = 'C:\\Users\\mundo\\Desktop\\Meus Estudos\\Altomações JS\\produtos_ficticios.xlsx';
function preencherCampo(x, y, valor) {
    // Mover o mouse para as coordenadas do campo
    robot.moveMouse(x, y);

    // Clicar no campo para selecioná-lo
    robot.mouseClick();

    // Colar o valor no campo
    clipboardy.writeSync(valor);
    robot.keyTap('v');
    // Pressionar a tecla Tab para sair do campo
    robot.keyTap('tab');
}
function abrirPlanilha() {
    console.log('Planilha já deve estar aberta antes de executar o código.');
}
const workbook = XLSX.readFile(planilha);
const sheet_produtos = workbook.Sheets[workbook.SheetNames[0]];
const dados = XLSX.utils.sheet_to_json(sheet_produtos);
dados.forEach((linha, index) => {
    setTimeout(() => {
        const nomeProduto = linha['Nome do produto'];
        const descricao = linha['Descrição'];
        const categoria = linha['Categoria'];
        const codigodoproduto = linha['Código do produto'];
        const peso = linha['Peso (em kg)'];
        const dimensoes = linha['Dimensões (L x A x P)'];
        const preco = linha['Preço'];
        const quantidadeemestoque = linha['Quantidade em estoque'];
        const dataDeValidade = linha['Data de validade'];
        const cor = linha['Cor'];
        const tamanho = linha['Tamanho'];
        const material = linha['Material'];
        const fabricante = linha['Fabricante'];
        const paisDeOrigem = linha['País de origem'];
        const observacoes = linha['Observações'];
        const codigoDeBarras = linha['Código de barras'];
        const localizacaoNoArmazem = linha['Localização no armazém'];

        // Colar os valores nas coordenadas especificadas
        preencherCampo(1158, 348, nomeProduto);
        preencherCampo(1150, 437, descricao);
        preencherCampo(1190, 567, categoria);
        preencherCampo(1167, 654, codigodoproduto);
        const pesoString = peso.toString();
        preencherCampo(1149, 737, pesoString);
        preencherCampo(1166, 826, dimensoes);
        while (index < dados.length - 1) {
            robot.moveMouse(1124, 881)
            robot.mouseClick()
            //
            preencherCampo(1166, 826, preco);
            preencherCampo(1166, 826, quantidadeemestoque);
            preencherCampo(1166, 826, dataDeValidade);
            preencherCampo(1166, 826, cor);
            if (tamanho == 'Pequeno') {
                robot.moveMouse(1134, 732);
                robot.mouseClick()
                robot.moveMouse(1134, 759)
                robot.mouseClick()
            } else if (tamanho == 'Médio') {
                robot.moveMouse(1134, 732);
                robot.mouseClick()
                robot.moveMouse(1134, 780)
                robot.mouseClick()
            } else if (tamanho == 'Grande') {
                robot.moveMouse(1134, 732);
                robot.mouseClick()
                robot.moveMouse(1134, 805)
                robot.mouseClick()
            }
            preencherCampo(1175, 818, material);
            robot.moveMouse(1124, 856)
            robot.mouseClick()
            preencherCampo(1165, 407, fabricante)
            preencherCampo(1143, 497, paisDeOrigem)
            preencherCampo(1137, 593, observacoes)
            preencherCampo(1139, 716, codigoDeBarras)
            preencherCampo(1138, 794, localizacaoNoArmazem)
            robot.moveMouse(1114, 849)
            robot.mouseClick() //concluir
            robot.moveMouse(1605, 193)
            robot.mouseClick() //OK
            return console.log("DEU ATÉ AQUI")
        }
    }, index * 3000);
});

abrirPlanilha(planilha);