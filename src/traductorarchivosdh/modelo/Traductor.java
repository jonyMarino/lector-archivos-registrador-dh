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
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.*;
import java.util.*;
import jxl.Workbook;
import jxl.write.DateFormat;
import jxl.write.Number;

import jxl.write.*;
import java.text.SimpleDateFormat;


public class Traductor {
    public static void traducirTxtAXls(LinkedList<Map<String,Object>> tabla,String pathXls){
        try
        {
          String filename = pathXls;
          WorkbookSettings ws = new WorkbookSettings();
          ws.setLocale(new Locale("en", "EN"));
          WritableWorkbook workbook = 
          Workbook.createWorkbook(new File(filename), ws);
          WritableSheet sheet = workbook.createSheet("Sheet1", 0);
          writeDataSheet(sheet,tabla);
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
      private static void writeDataSheet(WritableSheet sheet,LinkedList<Map<String,Object>> tabla) 
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
            
            
        Set<String> columnas = tabla.get(0).keySet();
        Iterator<String> itColumnas = columnas.iterator();
        int j = 0;
        while (itColumnas.hasNext()){            
            String nombreColumna = itColumnas.next();
            Label l = new Label(j,0,nombreColumna,cf);
            sheet.addCell(l);
            ++j;
        }

        int i=1;
        Iterator<Map<String, Object>> it = tabla.iterator();
        while (it.hasNext())
        {             
            j=0;
            Map<String, Object> map = it.next();
            itColumnas = columnas.iterator();
            while (itColumnas.hasNext()) {
                Object o = map.get(itColumnas.next());
                if(o == null)continue;
                String s = o.toString();
                Label l = new Label(j,i,s,cf);
                sheet.addCell(l);
                ++j;
            }
            ++i;
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
          Table[] tablaAccess ={db.getTable("Instrumento1"),db.getTable("Instrumento2"),db.getTable("Instrumento3"),db.getTable("Instrumento4"),db.getTable("Combinada")};
          Iterator<Map<String, Object>> it = tabla.iterator();
          SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
          int id=0;
        while (it.hasNext()) {
            Map<String, Object> map = it.next();
            String fecha = (String) map.get("Fecha");
            String hora = (String) map.get("Hora");           
            Date dFecha = formatter.parse(fecha);
            Object[] entrada = new Object[7*4];
            
            for(int j=0;j<4;j++){
                entrada[0+j*7] = map.get("VALOR_"+(j+1));
                entrada[1+j*7] = map.get("SP_"+(j+1));

                entrada[2+j*7] = map.get("POTENCIA_"+(j+1));
                entrada[3+j*7] = map.get("ALARMA_"+(j+1));
                if(entrada[3+j*7]==null){
                    entrada[3+j*7] = map.get("ALARMA1_"+(j+1));
                }
                entrada[4+j*7] = map.get("ALARMA2_"+(j+1));
                entrada[5+j*7] = map.get("ALARMA3_"+(j+1));
                entrada[6+j*7] = map.get("ALARMA4_"+(j+1));

                
                
                boolean empty=true;
                for(int i=0;i<7;i++){
                    if(entrada[j*7+i]==null) entrada[j*7+i] = "N.I.";
                    else empty = false;
                }
                
                //si es de un solo canal tiene otra nomenclatura
                if(empty && j==0){
                    entrada[0] = map.get("VALOR");

                    entrada[1] = map.get("SP");

                    entrada[2] = map.get("POTENCIA");
                    entrada[3] = map.get("ALARMA");
                    if(entrada[3]==null){
                        entrada[3] = map.get("ALARMA1");
                    }
                    entrada[4] = map.get("ALARMA2");
                    entrada[5] = map.get("ALARMA3");
                    entrada[6] = map.get("ALARMA4");
                    for(int i=0;i<7;i++){
                        if(entrada[i]==null) entrada[i] = "N.I.";
                        else empty = false;
                    }
                }
                if(!empty){
                    
                    tablaAccess[j].addRow(dFecha,hora,"",entrada[0+j*7],entrada[1+j*7],entrada[2+j*7],entrada[3+j*7],entrada[4+j*7],entrada[5+j*7],entrada[6+j*7]);
                }
            }
            //ahora que pase por los cuatro canales guardo en la combinada

            tablaAccess[4].addRow(id++,dFecha,hora,entrada[0],entrada[1],entrada[2],entrada[3],entrada[4],entrada[5],entrada[6],entrada[7],entrada[8],entrada[9],entrada[10],entrada[11],entrada[12],entrada[13],entrada[14],entrada[15],entrada[16],entrada[17],entrada[18],entrada[19],entrada[20],entrada[21],entrada[22],entrada[23],entrada[24],entrada[25],entrada[26],entrada[27]);

        }  
        db.close();
      }
      
      private static Integer onOffAPorcentaje(String s) throws Exception{
          if(s.equalsIgnoreCase("ON"))
            return 100;
          if(s.equalsIgnoreCase("OFF"))
            return 0;
          else{
              throw new Exception("Error de formateo");
          }
              
      }
      private static Float potenciaAFloat(String s){
          int indicePorciento=s.indexOf("%");
          String valor = s.substring(0, indicePorciento-1);
          return Float.parseFloat(valor);
      }
      public static Database crearBaseDeDatosDHSoftAccess(String pathMdb)throws IOException, SQLException{
        Database db = Database.create(new File(pathMdb),false);     
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
                    Map<String,Object> m = new Traductor.AdquisicionMap();
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
                           if(columna.matches("ALARMA.*"))
                               Valor = onOffAPorcentaje((String)Valor);
                           else if(columna.matches("POTENCIA.*"))
                               Valor = potenciaAFloat((String)Valor);
                           m.put(columnas.get(i+3), Valor);
                           ++i;
                       }
                       if(i!=0) //me fijo que no sea una linea vacia
                        lista.add(m);
                       
                       
                    }catch(Exception e){
                        //
                    }

                }
            }    
            return lista;
    }
        
    private static class AdquisicionMap extends LinkedHashMap<String,Object> implements Comparable<AdquisicionMap>{

        boolean equals(AdquisicionMap o) {
            String fecha1= (String)o.get("Fecha");
            String hora1= (String)o.get("Hora");
            String fecha2 = (String)get("Fecha");         
            String hora2 = (String)get("Hora");
            try {
                SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
                Date dFecha1 = f.parse(fecha1+" "+hora1);
                Date dFecha2 = f.parse(fecha2+" "+hora2);
                return dFecha2.equals(dFecha1);
            } catch (ParseException ex) {
                Logger.getLogger(Traductor.class.getName()).log(Level.SEVERE, null, ex);
            }
            return false;
        }
        @Override
        public boolean equals(Object o){
            return o instanceof AdquisicionMap && equals((AdquisicionMap)o);
        }

        @Override
        public int hashCode() {
            int hash = 3;
            return hash;
        }

        public int compareTo(AdquisicionMap o) {
            String fecha1= (String)o.get("Fecha");
            String hora1= (String)o.get("Hora");
            String fecha2 = (String)get("Fecha");         
            String hora2 = (String)get("Hora");
            try {
                SimpleDateFormat f = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
                Date dFecha1 = f.parse(fecha1+" "+hora1);
                Date dFecha2 = f.parse(fecha2+" "+hora2);
                return dFecha2.compareTo(dFecha1);
            } catch (ParseException ex) {
                Logger.getLogger(Traductor.class.getName()).log(Level.SEVERE, null, ex);
            }
            return 0;
        }
    
    }
    
}
