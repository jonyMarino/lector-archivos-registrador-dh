/*
 * TraductorArchivosDHView.java
 */

package traductorarchivosdh;

import java.awt.Toolkit;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.LinkedList;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.jdesktop.application.Action;
import org.jdesktop.application.ResourceMap;
import org.jdesktop.application.SingleFrameApplication;
import org.jdesktop.application.FrameView;
import org.jdesktop.application.TaskMonitor;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;
import javax.swing.Timer;
import javax.swing.Icon;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import traductorarchivosdh.modelo.Mixer;
import traductorarchivosdh.modelo.Traductor;


/**
 * The application's main frame.
 */
public class TraductorArchivosDHView extends FrameView {

    public TraductorArchivosDHView(SingleFrameApplication app) {
        super(app);

        //getFrame().setIconImage(Toolkit.getDefaultToolkit().getImage(ClassLoader.getSystemResource("resources\\Logo.png")));
        initComponents();
        // status bar initialization - message timeout, idle icon and busy animation, etc
        ResourceMap resourceMap = getResourceMap();
        int messageTimeout = resourceMap.getInteger("StatusBar.messageTimeout");
        messageTimer = new Timer(messageTimeout, new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                statusMessageLabel.setText("");
            }
        });
        messageTimer.setRepeats(false);
        int busyAnimationRate = resourceMap.getInteger("StatusBar.busyAnimationRate");
        for (int i = 0; i < busyIcons.length; i++) {
            busyIcons[i] = resourceMap.getIcon("StatusBar.busyIcons[" + i + "]");
        }
        busyIconTimer = new Timer(busyAnimationRate, new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                busyIconIndex = (busyIconIndex + 1) % busyIcons.length;
                statusAnimationLabel.setIcon(busyIcons[busyIconIndex]);
            }
        });
        idleIcon = resourceMap.getIcon("StatusBar.idleIcon");
        statusAnimationLabel.setIcon(idleIcon);
        progressBar.setVisible(false);

        // connecting action tasks to status bar via TaskMonitor
        TaskMonitor taskMonitor = new TaskMonitor(getApplication().getContext());
        taskMonitor.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                String propertyName = evt.getPropertyName();
                if ("started".equals(propertyName)) {
                    if (!busyIconTimer.isRunning()) {
                        statusAnimationLabel.setIcon(busyIcons[0]);
                        busyIconIndex = 0;
                        busyIconTimer.start();
                    }
                    getProgressBar().setVisible(true);
                    getProgressBar().setIndeterminate(true);
                } else if ("done".equals(propertyName)) {
                    busyIconTimer.stop();
                    statusAnimationLabel.setIcon(idleIcon);
                    getProgressBar().setVisible(false);
                    getProgressBar().setValue(0);
                } else if ("message".equals(propertyName)) {
                    String text = (String)(evt.getNewValue());
                    statusMessageLabel.setText((text == null) ? "" : text);
                    messageTimer.restart();
                } else if ("progress".equals(propertyName)) {
                    int value = (Integer)(evt.getNewValue());
                    getProgressBar().setVisible(true);
                    getProgressBar().setIndeterminate(false);
                    getProgressBar().setValue(value);
                }
            }
        });
        
    }

    @Action
    public void showAboutBox() {
        if (aboutBox == null) {
            JFrame mainFrame = TraductorArchivosDHApp.getApplication().getMainFrame();
            aboutBox = new TraductorArchivosDHAboutBox(mainFrame);
            aboutBox.setLocationRelativeTo(mainFrame);
        }
        TraductorArchivosDHApp.getApplication().show(aboutBox);
    }

    /** This method is called from within the constructor to
     * initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is
     * always regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        mainPanel = new javax.swing.JPanel();
        jButton1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jPanel1 = new javax.swing.JPanel();
        jButton3 = new javax.swing.JButton();
        jTextField1 = new javax.swing.JTextField();
        menuBar = new javax.swing.JMenuBar();
        javax.swing.JMenu fileMenu = new javax.swing.JMenu();
        javax.swing.JMenuItem exitMenuItem = new javax.swing.JMenuItem();
        javax.swing.JMenu helpMenu = new javax.swing.JMenu();
        javax.swing.JMenuItem aboutMenuItem = new javax.swing.JMenuItem();
        statusPanel = new javax.swing.JPanel();
        javax.swing.JSeparator statusPanelSeparator = new javax.swing.JSeparator();
        statusMessageLabel = new javax.swing.JLabel();
        statusAnimationLabel = new javax.swing.JLabel();
        progressBar = new javax.swing.JProgressBar();

        mainPanel.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        mainPanel.setName("mainPanel"); // NOI18N
        mainPanel.setPreferredSize(new java.awt.Dimension(421, 250));

        javax.swing.ActionMap actionMap = org.jdesktop.application.Application.getInstance(traductorarchivosdh.TraductorArchivosDHApp.class).getContext().getActionMap(TraductorArchivosDHView.class, this);
        jButton1.setAction(actionMap.get("traducir")); // NOI18N
        org.jdesktop.application.ResourceMap resourceMap = org.jdesktop.application.Application.getInstance(traductorarchivosdh.TraductorArchivosDHApp.class).getContext().getResourceMap(TraductorArchivosDHView.class);
        jButton1.setActionCommand(resourceMap.getString("jButton1.actionCommand")); // NOI18N
        jButton1.setAlignmentY(0.0F);
        jButton1.setLabel(resourceMap.getString("jButton1.label")); // NOI18N
        jButton1.setName("jButton1"); // NOI18N

        jButton2.setAction(actionMap.get("traducirParaDHSoft")); // NOI18N
        jButton2.setLabel(resourceMap.getString("jButton2.label")); // NOI18N
        jButton2.setName("jButton2"); // NOI18N

        jPanel1.setBorder(javax.swing.BorderFactory.createTitledBorder(resourceMap.getString("jPanel1.border.title"))); // NOI18N
        jPanel1.setName("jPanel1"); // NOI18N

        jButton3.setAction(actionMap.get("seleccionarArchivo")); // NOI18N
        jButton3.setLabel(resourceMap.getString("jButton3.label")); // NOI18N
        jButton3.setName("jButton3"); // NOI18N

        jTextField1.setText(resourceMap.getString("jTextField1.text")); // NOI18N
        jTextField1.setEnabled(false);
        jTextField1.setName("jTextField1"); // NOI18N

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jTextField1, javax.swing.GroupLayout.DEFAULT_SIZE, 272, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton3)
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton3, javax.swing.GroupLayout.PREFERRED_SIZE, 20, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout mainPanelLayout = new javax.swing.GroupLayout(mainPanel);
        mainPanel.setLayout(mainPanelLayout);
        mainPanelLayout.setHorizontalGroup(
            mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(mainPanelLayout.createSequentialGroup()
                .addContainerGap()
                .addGroup(mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
                    .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 256, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 256, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap(14, Short.MAX_VALUE))
        );
        mainPanelLayout.setVerticalGroup(
            mainPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(mainPanelLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton1)
                .addGap(21, 21, 21)
                .addComponent(jButton2)
                .addContainerGap(21, Short.MAX_VALUE))
        );

        menuBar.setName("menuBar"); // NOI18N

        fileMenu.setLabel(resourceMap.getString("fileMenu.label")); // NOI18N
        fileMenu.setName("fileMenu"); // NOI18N

        exitMenuItem.setAction(actionMap.get("quit")); // NOI18N
        exitMenuItem.setLabel(resourceMap.getString("exitMenuItem.label")); // NOI18N
        exitMenuItem.setName("exitMenuItem"); // NOI18N
        fileMenu.add(exitMenuItem);

        menuBar.add(fileMenu);

        helpMenu.setLabel(resourceMap.getString("helpMenu.label")); // NOI18N
        helpMenu.setName("helpMenu"); // NOI18N

        aboutMenuItem.setAction(actionMap.get("showAboutBox")); // NOI18N
        aboutMenuItem.setLabel(resourceMap.getString("aboutMenuItem.label")); // NOI18N
        aboutMenuItem.setName("aboutMenuItem"); // NOI18N
        helpMenu.add(aboutMenuItem);

        menuBar.add(helpMenu);

        statusPanel.setName("statusPanel"); // NOI18N

        statusPanelSeparator.setName("statusPanelSeparator"); // NOI18N

        statusMessageLabel.setName("statusMessageLabel"); // NOI18N

        statusAnimationLabel.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
        statusAnimationLabel.setName("statusAnimationLabel"); // NOI18N

        progressBar.setMaximum(1000);
        progressBar.setName("progressBar"); // NOI18N

        javax.swing.GroupLayout statusPanelLayout = new javax.swing.GroupLayout(statusPanel);
        statusPanel.setLayout(statusPanelLayout);
        statusPanelLayout.setHorizontalGroup(
            statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(statusPanelLayout.createSequentialGroup()
                .addContainerGap()
                .addComponent(statusMessageLabel)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 253, Short.MAX_VALUE)
                .addComponent(progressBar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(statusAnimationLabel)
                .addContainerGap())
            .addComponent(statusPanelSeparator, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 423, Short.MAX_VALUE)
        );
        statusPanelLayout.setVerticalGroup(
            statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, statusPanelLayout.createSequentialGroup()
                .addComponent(statusPanelSeparator, javax.swing.GroupLayout.PREFERRED_SIZE, 2, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addGroup(statusPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(statusMessageLabel)
                    .addComponent(statusAnimationLabel)
                    .addComponent(progressBar, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(3, 3, 3))
        );

        setComponent(mainPanel);
        setMenuBar(menuBar);
        setStatusBar(statusPanel);
    }// </editor-fold>//GEN-END:initComponents

    @Action
    public void traducir() throws FileNotFoundException, Exception {
        
        if(files == null){
            JOptionPane.showMessageDialog(null,"Debe seleccionar al menos un archivo del registrador.");
            return;
        }
        
        String archivoXls;
        JFileChooser fc = new JFileChooser();
        fc.addChoosableFileFilter(new ExcelFilter());
        int returnVal =fc.showSaveDialog(jPanel1);
     
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            archivoXls = file.getAbsolutePath();
            if(!Utils.getExtension(file).equals("mdb"))
                archivoXls+=".xls";
        } else {
            return;
        }
        getFrame().setEnabled(false);
        final LinkedList<Map<String, Object>> lista = Mixer.getTablaCombinada(files);
        Traductor.traducirTxtAXls(lista, archivoXls);
        JOptionPane.showMessageDialog(null,"Ha creado exitosamente el archivo: "+archivoXls);
        getFrame().setEnabled(true);
    }

    /**
     * @return the progressBar
     */
    public javax.swing.JProgressBar getProgressBar() {
        return progressBar;
    }

    /**
     * @param progressBar the progressBar to set
     */
    public void setProgressBar(javax.swing.JProgressBar progressBar) {
        this.progressBar = progressBar;
    }
    
    class ExcelFilter extends javax.swing.filechooser.FileFilter {
        public boolean accept(File file) {
            String filename = file.getName();
            return filename.endsWith(".xls");
        }
        public String getDescription() {
            return "*.xls";
        }
    }


    @Action
    public void seleccionarArchivo() {
        JFileChooser fc = new JFileChooser();
        fc.setCurrentDirectory(new File(System.getProperty("user.dir")));
        fc.setMultiSelectionEnabled(true);
        int returnVal = fc.showOpenDialog(jPanel1);
        if (returnVal == JFileChooser.APPROVE_OPTION) {
          files=fc.getSelectedFiles();
          /*files = new File[1];
          files[0]=fc.getSelectedFile();*/
          String sFiles=new String();
          for(int i=0;i<files.length;i++){
              sFiles += files[i].getName()+"; ";
          }
          jTextField1.setText(sFiles);  

        } else {
            System.out.println("File access cancelled by user.");
        }
        fc.setVisible(false);
        //  File f = jFileChooser1.getSelectedFile();
        //  jTextField1.setText(f.toString());
    }
    

    @Action
    public void traducirParaDHSoft() throws IOException, SQLException, FileNotFoundException, Exception {

        
        if(files == null){
            JOptionPane.showMessageDialog(null,"Debe seleccionar al menos un archivo del registrador.");
            return;
        }
        
        String archivoMdb;
        JFileChooser fc = new JFileChooser();
        fc.addChoosableFileFilter(new AccessFilter());
        int returnVal =fc.showSaveDialog(jPanel1);
        
        
        if (returnVal == JFileChooser.APPROVE_OPTION) {
            File file = fc.getSelectedFile();
            archivoMdb = file.getAbsolutePath();
            if(!Utils.getExtension(file).equals("mdb"))
                archivoMdb+=".mdb";

        } else {
            return;
        }
        
        try{
            //final LinkedList<Map<String, Object>> lista = Traductor.getTabla(files[0].getAbsolutePath());
            final LinkedList<Map<String, Object>> lista = Mixer.getTablaCombinada(files);
            getFrame().setEnabled(false);
            Traductor.traducirTxtADHSoft(lista, archivoMdb);
            JOptionPane.showMessageDialog(null,"Ha creado exitosamente el archivo: "+archivoMdb);
            getFrame().setEnabled(true);
        }catch(Exception e){
            getFrame().setEnabled(false);
            e.getCause();
        }

    }
    
    class AccessFilter extends javax.swing.filechooser.FileFilter {
        public boolean accept(File file) {
            String filename = file.getName();
            return filename.endsWith(".mdb");
        }
        public String getDescription() {
            return "*.mdb";
        }
    }
    
    

    @Action
    public void graficar() throws FileNotFoundException, IOException, Exception {
        
    
        
        if(files == null){
            JOptionPane.showMessageDialog(null,"Debe seleccionar al menos un archivo del registrador.");
            return;
        }

        //final LinkedList<Map<String, Object>> lista = Traductor.getTabla(files[0].getAbsolutePath());
        final LinkedList<Map<String, Object>> lista = Mixer.getTablaCombinada(files);
        java.awt.EventQueue.invokeLater(new Runnable() {

            public void run() {
                try {
                    new Graficador(lista).setVisible(true);
                } catch (ParseException ex) {
                    Logger.getLogger(TraductorArchivosDHView.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        });
        
         
    }
    private File[] files=null;




    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JPanel mainPanel;
    private javax.swing.JMenuBar menuBar;
    private javax.swing.JProgressBar progressBar;
    private javax.swing.JLabel statusAnimationLabel;
    private javax.swing.JLabel statusMessageLabel;
    private javax.swing.JPanel statusPanel;
    // End of variables declaration//GEN-END:variables

    private final Timer messageTimer;
    private final Timer busyIconTimer;
    private final Icon idleIcon;
    private final Icon[] busyIcons = new Icon[15];
    private int busyIconIndex = 0;

    private JDialog aboutBox;
}
class Utils {
    /*
     * Get the extension of a file.
     */  
    public static String getExtension(File f) {
        String ext = "";
        String s = f.getName();
        int i = s.lastIndexOf('.');

        if (i > 0 &&  i < s.length() - 1) {
            ext = s.substring(i+1).toLowerCase();
        }
        return ext;
    }
}
