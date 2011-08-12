/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package traductorarchivosdh.modelo;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author rbravo
 */
public class TraductorTest {
    
    public TraductorTest() {
    }

    @BeforeClass
    public static void setUpClass() throws Exception {
    }

    @AfterClass
    public static void tearDownClass() throws Exception {
    }
    
    @Before
    public void setUp() {
    }
    
    @After
    public void tearDown() {
    }

    /**
     * Test of traducirTxtAXls method, of class Traductor.
     */
    @Test
    public void testTraducirTxtAXls() {
        System.out.println("traducirTxtAXls");
        for(int i=1;i<7;i++){
            String pathTxt = System.getProperty("user.dir")+"\\test"+i+".txt";
            String pathXls = "test"+i+".xls";
            Traductor.traducirTxtAXls(pathTxt, pathXls);
            System.out.println("traducirTxtAXls test"+i);
        }
        // TODO review the generated test code and remove the default call to fail.

    }

    /**
     * Test of traducirTxtADHSoft method, of class Traductor.
     */
    @Test
    public void testTraducirTxtADHSoft() throws Exception {
        System.out.println("traducirTxtADHSoft");
        for(int i=1;i<7;i++){
            String pathTxt = System.getProperty("user.dir")+"\\test"+i+".txt";
            String pathMdb = "test"+i+".mdb";
            Traductor.traducirTxtADHSoft(pathTxt, pathMdb);
            System.out.println("traducirTxtADHSoft test"+i);
        }
        // TODO review the generated test code and remove the default call to fail.

    }
}
