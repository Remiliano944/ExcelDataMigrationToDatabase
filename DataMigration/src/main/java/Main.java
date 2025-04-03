import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.*;

public class Main {

    public static void main(String[] args) {
        //bloco try para detectar erros
        try{
            InputStream file = new FileInputStream("src/main/java/data.xlsx");
            //instanciar o workbook usando o elemento obtido do FileInputStream. Isso implica que o programa simplesmente traduz com base na dedução do que o objeto da classe FileInputStream possui? Se sim. Wow.

           // XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Vou tentar implementar por meio do WorkbookFactory. Já que, segundo a documentação, ele já cuida de tratar o tipo do arquivo.
            //código quebra quando chega nessa parte. "NoClassDefFOundError" "ClassNotFoundException".
            //aparentemente é falta de dependências. Oh... Agora sim o uso do maven faz sentido!! Agora eu entendi!!
            Workbook workbook = WorkbookFactory.create(file);

            //declarando FormulaEvaluator para poder trabalhar com formulas em Excel.
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            //O workbook retorna um objeto da classe CreationHelper, que possui em seus métodos a capacidade de retornar um objeto da classe FormulaEvaluator

            //Agora sim vamos para as tabelinhas.

            Sheet sheet = workbook.getSheetAt(0);
            //Melhor ainda. Para cada dado que existir no limite da coleção.
            for (Row row : sheet) {

                //uma vez extraido os dados da coluna. Vamos para as linhas afim de obter suas informações.

                Iterator<Cell> cellIterator = row.cellIterator();

                //enquanto houver dados ou conteúdo na linha. Continue :)

                while (cellIterator.hasNext()) {

                    //obtendo os dados da linha

                    Cell cell = cellIterator.next();
                    //como o programa pode obter vários tipos de informações (sendo estas numéricas ou alfabéticas)...

                    //Aplicação do Evaluator aqui. Considerando que a classe já lida com qualquer tipo de valor...
                    //Sua implementação mais cedo há de beneficiar melhor o meu projeto, pois assim poderei usar o default case.

                    CellValue cellValue = evaluator.evaluate(cell);

                    switch (cellValue.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + ", ");
                            break;

                        case NUMERIC:
                            int valor = (int)cellValue.getNumberValue();
                            //captar a largura de digitos do valor
                            int length = String.valueOf(valor).length();

                            //rodar em looping para que enquanto a largura for maior do que o limite permitido (neste caso 7) dividir o número por 10.
                            while(length > 7){
                                valor /= 10;
                                //O numero perde uma casa decimal, e por ser um inteiro, a informação "perdida" não é contada, oque realmente implica que o número perdeu o digito
                                length--;
                            }
                            System.out.println(valor + ", ");
                            break;

                        case ERROR:
                            System.out.println("Erro, valor nulo ou fora de cogitação.");
                            break;
                    }
                }
                System.out.println(" ");
            }
            file.close();
            workbook.close();
        }
        catch (Exception e ){
            //coletar a exceção para que eu possa lidar com ela mais tarde.
            e.printStackTrace();
        }
    }
}