/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package traductorarchivosdh.modelo;

/**
 *
 * @author jony
 */
import com.healthmarketscience.jackcess.ColumnBuilder;
import com.healthmarketscience.jackcess.DataType;
import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.Table;
import com.healthmarketscience.jackcess.TableBuilder;
import java.io.*;
import java.nio.channels.FileChannel;
import java.sql.SQLException;
import java.sql.Types;
import java.text.ParseException;
import jxl.*;
import java.util.*;
import jxl.Workbook;
import jxl.write.DateFormat;
import jxl.write.Number;

import jxl.write.*;
import java.text.SimpleDateFormat;


public class Traductor {
    public static void traducirTxtAXls(String pathTxt,String pathXls){
        try
        {
          String filename = pathXls;
          WorkbookSettings ws = new WorkbookSettings();
          ws.setLocale(new Locale("en", "EN"));
          WritableWorkbook workbook = 
          Workbook.createWorkbook(new File(filename), ws);
          WritableSheet s = workbook.createSheet("Sheet1", 0);
          writeDataSheet(s,pathTxt);
          workbook.write();
          workbook.close();      
        }
        catch (IOException e)
        {
          e.printStackTrace();
        }
        catch (WriteException e)
        {
          e.printStackTrace();
        }

    }
      private static void writeDataSheet(WritableSheet s,String pathTxt) 
        throws WriteException
      {


        /* Format the Font */
        WritableFont wf = new WritableFont(WritableFont.ARIAL, 
          10, WritableFont.BOLD);
        WritableCellFormat cf = new WritableCellFormat(wf);
        cf.setWrap(true);

        /* Creates Label and writes date to one cell of sheet*/
        
        WritableCellFormat cf1 = 
          new WritableCellFormat(DateFormats.FORMAT9);

        /*DateTime dt = 
          new DateTime(0,1,new Date(), cf1, DateTime.GMT);

        s.addCell(dt);
        */
        
        try
        {
            BufferedReader reader = new BufferedReader(new FileReader(pathTxt));
            int i=0;
            while (reader.ready())
            {             
                int j=0;
                String line = reader.readLine();
                StringTokenizer st = new StringTokenizer(line);
                while (st.hasMoreTokens()) {
                    Label l = new Label(j,i,st.nextToken(),cf);
                    s.addCell(l);
                    ++j;
                }
                ++i;
            }
            reader.close();
        }
        catch (Exception e)
        {
            System.err.format("Exception occurred trying to read '%s'.", pathTxt);
            e.printStackTrace();
        }

     }
/*
      public void traducirTxtADHSoft(String pathTxt,String pathMdb) throws IOException, SQLException{
        Database db = Database.create(new File("new.mdb"));
        Table newTable = new TableBuilder("NewTable")
          .addColumn(new ColumnBuilder("a")
                     .setSQLType(Types.INTEGER)
                     .toColumn())
          .addColumn(new ColumnBuilder("b")
                     .setSQLType(Types.VARCHAR)
                     .toColumn())
          .toTable(db);
        newTable.addRow(1, "foo");

        }
      */
      
      public static void traducirTxtADHSoft(LinkedList<Map<String,Object>> tabla,String pathMdb) throws IOException, SQLException, ParseException{
          Database db = crearBaseDeDatosDHSoftAccess(pathMdb);
          Table tablaAccess = db.getTable("Instrumento1");
          Iterator<Map<String, Object>> it = tabla.iterator();
        while (it.hasNext()) {
            Map<String, Object> map = it.next();
            String fecha = (String) map.get("Fecha");
            String hora = (String) map.get("Hora");
            SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy"); 
            Date dFecha = formatter.parse(fecha);

            Object alarma = map.get("ALARMA");
            if(alarma==null){
                alarma = map.get("ALARMA1");
            }
            tablaAccess.addRow(dFecha,hora,"",map.get("VALOR"),map.get("SP"),map.get("POTENCIA"),alarma,map.get("ALARMA2"),map.get("ALARMA3"),map.get("ALARMA4")); 

        }
          
      }
      
      private static int onOffAPorcentaje(String s) throws Exception{
          if(s.equalsIgnoreCase("ON"))
            return 100;
          if(s.equalsIgnoreCase("OFF"))
            return 0;
          else{
              throw new Exception("Error de formateo");
          }
              
      }
      private static float potenciaAFloat(String s){
          int indicePorciento=s.indexOf("%");
          String valor = s.substring(0, indicePorciento-1);
          return Float.parseFloat(valor);
      }
      public static Database crearBaseDeDatosDHSoftAccess(String pathMdb)throws IOException, SQLException{
        Database db = Database.create(new File(pathMdb));
        for(int i=1;i<21;i++){
            TableBuilder instrumento = new TableBuilder("Instrumento"+i)
            .addColumn(new ColumnBuilder("Fecha")
             .setSQLType(Types.DATE)
             .setType(DataType.SHORT_DATE_TIME)
             .toColumn())
            .addColumn(new ColumnBuilder("Hora")
             .setSQLType(Types.VARCHAR)       
             .setLength(18)
             .toColumn())
            .addColumn(new ColumnBuilder("Partida")
             .setSQLType(Types.VARCHAR)
             .setLength(24)
             .toColumn());
            String columnasInstrumento[]={"Valor_medido","Set_point","Potencia","Alarma_1","Alarma_2","Alarma_3","Alarma_4",null};
            for(int j=0;columnasInstrumento[j]!=null;j++){
                instrumento.addColumn(new ColumnBuilder(columnasInstrumento[j])
                    .setSQLType(Types.VARCHAR)
                    .setLength(10)
                    .toColumn());
            }         
            instrumento.toTable(db);
            //newTable.getColumn("Potencia").write(1,0);
          //  newTable.addRow("19/01/1987");
        }
        TableBuilder combinada = new TableBuilder("Combinada")
            .addColumn(new ColumnBuilder("Id")
             .setSQLType(Types.INTEGER)
             .setAutoNumber(true)
             .toColumn())
            .addColumn(new ColumnBuilder("Fecha")
             .setSQLType(Types.DATE)
             .setType(DataType.SHORT_DATE_TIME)   
             .toColumn())
            .addColumn(new ColumnBuilder("Hora")
             .setSQLType(Types.VARCHAR)       
             .setLength(18)
             .toColumn());
        String columnas[]={"V_medido","SP","Pot","AL1","AL2","AL3","AL4",null};
        for(int i=1;i<21;i++){
            for(int j=0;columnas[j]!=null;j++){
                combinada.addColumn(new ColumnBuilder(columnas[j]+"(Ins"+i+")")
                    .setSQLType(Types.VARCHAR)
                    .setLength(10)
                    .toColumn());
            }
         }
       combinada.toTable(db); 
       return db;
      }
      
      public static void copyFile(File sourceFile, File destFile) throws IOException {  
          if(!destFile.exists()) {   
              destFile.createNewFile();
          }   
          FileChannel source = null;  
          FileChannel destination = null;  
          try {   
              source = new FileInputStream(sourceFile).getChannel();      
               destination = new FileOutputStream(destFile).getChannel();   
               destination.transferFrom(source, 0, source.size());  
          }  
          finally {   
              if(source != null) {    
                source.close();
              }   
          }   
          if(destination != null) {    
              destination.close();   } 
      }

    private static LinkedList<String> obtenerColumnas(String line) {
        StringTokenizer st = new StringTokenizer(line);
        LinkedList<String> result = new LinkedList<String>();
        while (st.hasMoreTokens()) {
           result.add(st.nextToken());
        }
        return result;
    }
    
        public static LinkedList<Map<String,Object>>  getTabla(String pathTxt) throws FileNotFoundException, IOException, Exception{
            BufferedReader reader = new BufferedReader(new FileReader(pathTxt));
           
            String line;
            do{
                line = reader.readLine();
            }while(line.isEmpty()/*&& reader.ready()*/);
            if(!reader.ready())
                throw new Exception("El archivo no contiene informaci"+162+"n");
            LinkedList<String> columnas = obtenerColumnas(line);
           // LinkedHashMap<String,Object> hash = new LinkedHashMap<String,Object>();
           // hash.
            
            LinkedList<Map<String,Object>> lista = new LinkedList<Map<String,Object>>();
            
            while (reader.ready())
            {  
                line = reader.readLine();
                StringTokenizer st = new StringTokenizer(line);

                if (st.hasMoreTokens()) {
                    Map<String,Object> m = new TreeMap<String,Object>();
                    try{
                       String id=st.nextToken();
                       String fecha=st.nextToken();
                       //java.text.DateFormat df = java.text.DateFormat.getDateInstance(java.text.DateFormat.SHORT);

                       //Date dFecha = df.parse(fecha);
                       m.put("Fecha", fecha);
                       String hora=st.nextToken();
                       m.put("Hora", hora);
                       
                       int i=0;
                       while(st.hasMoreTokens()){
                           String columna = columnas.get(i+3);
                           Object Valor=st.nextToken();
                           if(columna.length()>5 && columna.substring(0, 6).equals("ALARMA"))
                               Valor = onOffAPorcentaje((String)Valor);
                           else if(columna.equals("POTENCIA"))
                               Valor = potenciaAFloat((String)Valor);
                           m.put(columnas.get(i+3), Valor);
                           ++i;
                       }

                       Object alarma = m.get("ALARMA");
                       if(alarma==null){
                           alarma = m.get("ALARMA1");
                       }
                       lista.add(m);
                       
                       
                    }catch(NoSuchElementException e){
                        throw e;
                    }

                }
            }    
            return lista;
    }
    
}
