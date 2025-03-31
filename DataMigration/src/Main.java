import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.*;

public class Main {

    public static void main(String[] args) {
        //bloco try para detectar erros
        try{
            InputStream file = new FileInputStream("src/data.xlsx");
            //instanciar o workbook usando o elemento obtido do FileInputStream. Isso implica que o programa simplesmente traduz com base na dedução do que o objeto da classe FileInputStream possui? Se sim. Wow.

           // XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Vou tentar implementar por meio do WorkbookFactory. Já que, segundo a documentação, ele já cuida de tratar o tipo do arquivo.
            //código quebra quando chega nessa parte. "NoClassDefFOundError" "ClassNotFoundException".
            //aparentemente é falta de dependências. Oh... Agora sim o uso do maven faz sentido!! Agora eu entendi!!
            Workbook workbook = WorkbookFactory.create(file);

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


                    switch (cell.getCellType()) {
                        //Caso sejam letras/alfabético
                        case STRING:
                            System.out.print(cell.getStringCellValue() + ", ");
                            break;
                        case NUMERIC:
                            System.out.print((int)cell.getNumericCellValue() + ", ");
                            break;
                        default:
                            //para fazer forma de ler formulas de excel.
                           CellValue evaluate();
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