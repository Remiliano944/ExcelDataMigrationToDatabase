import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) throws Exception {
        //bloco try para detectar erros
        try{
            FileInputStream file = new FileInputStream(
                //por enquanto vou apenas considerar que o arquivo está em um lugar fixo. Em breve ei de implementar uma interface com busca para que o usuário selecione o arquivo.
                new File("../../data.xlsx"));
            //instanciar o workbook usando o elemento obtido do FileInputStream. Isso implica que o programa simplesmente traduz com base na dedução do que o objeto da classe FileInputStream possui? Se sim. Wow.
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Agora sim vamos para as tabelinhas.
            XSSFSheet sheet = workbook.getSheetAt(0);

            //iterar em cada tabela. Um por um
            Iterator<Row> rowIterator = sheet.iterator();

            //vamos implementar um loop while para simplificar.
            //Enquanto houver dados ou conteúdo na coluna. Continue :)
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();

                //uma vez extraido os dados da coluna. Vamos para as linhas afim de obter suas informações.
                Iterator<Cell> cellIterator = row.cellIterator();

                //enquanto houver dados ou conteúdo na linha. Continue :)
                while(cellIterator.hasNext()){

                    //obtendo os dados da linha
                    Cell cell = cellIterator.next();

                    //como o programa pode obter vários tipos de informações (sendo estas numéricas ou alfabéticas)...
                    switch(cell.getCellType()) {

                        //Caso sejam letras/alfabético
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "t");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                    }
                }
                System.out.println(" ");
            }
            file.close();
        }
        catch (Exception e ){
            //coletar a exceção para que eu possa lidar com ela mais tarde.
            e.printStackTrace();
        }
    }
}