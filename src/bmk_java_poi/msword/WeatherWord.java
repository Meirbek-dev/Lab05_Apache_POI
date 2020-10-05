package bmk_java_poi.msword;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;

public class WeatherWord extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread {

        @Override
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "weather_template.doc")) {
                doc = new HWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            try {
                doc.getRange().replaceText("$Город", jTextField_city.getText());
                doc.getRange().replaceText("$День", jTextField_day.getText());
            } catch (Exception ex) {
                System.err.println("Error replaceText!");
            }

            try (FileOutputStream fos = new FileOutputStream(dir + "weather.doc")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                Desktop.getDesktop().open(new File(dir + "weather.doc"));
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    public WeatherWord() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField_city = new javax.swing.JTextField();
        jTextField_day = new javax.swing.JTextField();
        jButton_write = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Прогноз погоды в MS Word");
        getContentPane().setLayout(null);
        getContentPane().add(jTextField_city);
        jTextField_city.setBounds(135, 37, 110, 33);
        getContentPane().add(jTextField_day);
        jTextField_day.setBounds(135, 70, 110, 33);

        jButton_write.setText("Записать в Word");
        jButton_write.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_writeActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_write);
        jButton_write.setBounds(260, 335, 130, 23);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/bmk_java_poi/msword/forecast.png"))); // NOI18N
        getContentPane().add(jLabel2);
        jLabel2.setBounds(0, 0, 640, 360);

        setBounds(0, 0, 654, 396);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_writeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_writeActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton_writeActionPerformed

 public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(WeatherWord.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }

        java.awt.EventQueue.invokeLater(() -> {
            new WeatherWord().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_write;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_city;
    private javax.swing.JTextField jTextField_day;
    // End of variables declaration//GEN-END:variables
}
