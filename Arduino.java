
import java.io.BufferedReader;
import java.io.InputStreamReader;
import gnu.io.CommPortIdentifier; 
import gnu.io.SerialPort;
import gnu.io.SerialPortEvent; 
import gnu.io.SerialPortEventListener; 
import java.io.File;
import java.io.FileOutputStream;
import java.util.Enumeration;


//EXCEL
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

//REGLA EXCEL
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;

//FECHA
import java.util.Calendar;
import javax.swing.JFileChooser;

//audio
import java.io.*;
import sun.audio.*;




public class Arduino extends javax.swing.JFrame implements SerialPortEventListener {
    
    
    int aux=0;
    int hora []= new int [1000];
    int minutos []= new int [1000];
    int segundos []= new int [1000];
    double temperatura []= new double [1000];
    
    
    SerialPort serialPort;
        /** The port we're normally going to use. */
	private static final String PORT_NAMES[] = { 
			"/dev/tty.usbserial-A9007UX1", // Mac OS X
                        "/dev/ttyACM0", // Raspberry Pi
			"/dev/ttyUSB0", // Linux
			"COM4", // Windows
	};

    private BufferedReader input;
    //private OutputStream output;
    private static final int TIME_OUT = 2000;
    private static final int DATA_RATE = 9600;
    
    


    public Arduino() {
        initComponents();
        this.setTitle("Registrador de Temperatura");
    }

    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">                          
    private void initComponents() {

        jFileChooser1 = new javax.swing.JFileChooser();
        jFileChooser2 = new javax.swing.JFileChooser();
        jButton1 = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTextArea1 = new javax.swing.JTextArea();
        jButton3 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        jComboBox1 = new javax.swing.JComboBox();
        jLabel1 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jButton1.setText("Adquirir");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        jTextArea1.setColumns(20);
        jTextArea1.setRows(5);
        jTextArea1.setMinimumSize(new java.awt.Dimension(3, 22));
        jScrollPane1.setViewportView(jTextArea1);

        jButton3.setText("Generar Archivo");
        jButton3.setEnabled(false);
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jButton2.setText("Detener");
        jButton2.setEnabled(false);
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "", "" }));
        jComboBox1.setEnabled(false);
        jComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox1ActionPerformed(evt);
            }
        });

        jLabel1.setFont(new java.awt.Font("Tahoma", 0, 8)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(51, 51, 51));
        jLabel1.setText("Versión 1.0");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(40, 40, 40)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(0, 0, Short.MAX_VALUE)
                        .addComponent(jButton3)
                        .addGap(81, 81, 81)
                        .addComponent(jLabel1)
                        .addContainerGap())
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jButton1)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 70, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(54, 54, 54)
                                .addComponent(jButton2))
                            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 319, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addContainerGap(36, Short.MAX_VALUE))))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton1)
                    .addComponent(jButton2)
                    .addComponent(jComboBox1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jButton3)
                    .addComponent(jLabel1))
                .addGap(0, 8, Short.MAX_VALUE))
        );

        pack();
    }// </editor-fold>                        

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {                                         
        // TODO add your handling code here:
      
        
        
          jButton1.setEnabled(false);
        jButton2.setEnabled(false);
        jButton3.setEnabled(false);
        jComboBox1.setEnabled(true);

        java.util.Enumeration<CommPortIdentifier> portEnum = CommPortIdentifier.getPortIdentifiers();
        
        int i = 0;
        String[] r = new String[5];

        while (portEnum.hasMoreElements() && i < 5) {
            CommPortIdentifier portIdentifier = portEnum.nextElement();
            r[i] = portIdentifier.getName();
            i++;
        }
        
         jComboBox1.setModel(new javax.swing.DefaultComboBoxModel(r));
         
          
               
    }                                        

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {                                         
        // TODO add your handling code here:
        int i;
        jButton1.setEnabled(true);
        jButton2.setEnabled(false);
        jButton3.setEnabled(false);
        jComboBox1.setEnabled(false);
        
        
       
       Workbook workbook = new HSSFWorkbook();

                Sheet sheet = workbook.createSheet("Temperatura");

                Cell celda;
                Cell celda1;

                celda = sheet.createRow(0).createCell(0);
                celda.setCellValue("HORA");
                celda1 = sheet.getRow(0).createCell(1);
                celda1.setCellValue("TEMPERATURA");
                
                for (i = 1; i < aux; i++) {
                    celda = sheet.createRow(i).createCell(1);
                    celda.setCellValue(temperatura[i]);
                    celda1 = sheet.getRow(i).createCell(0);
                    celda1.setCellValue(hora[i]+":"+minutos[i]+":"+segundos[i]);
                }
                
                
                 
                //create conditional formats
                SheetConditionalFormatting cond_formato = sheet.getSheetConditionalFormatting();
                
                //create rules
                ConditionalFormattingRule rule = cond_formato.createConditionalFormattingRule(ComparisonOperator.GT,"28");
                ConditionalFormattingRule rule1 = cond_formato.createConditionalFormattingRule(ComparisonOperator.LT,"18");
                
                //change background color
                
                PatternFormatting background = rule.createPatternFormatting();
                background.setFillBackgroundColor(IndexedColors.ORANGE.getIndex());
                
                PatternFormatting background1 = rule1.createPatternFormatting();
                background1.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
                
                CellRangeAddress [] range ={CellRangeAddress.valueOf("B2:B"+aux)};
                
                cond_formato.addConditionalFormatting(range,rule);
                cond_formato.addConditionalFormatting(range,rule1);
                

                String nombre="";
                JFileChooser file=new JFileChooser();
                file.showSaveDialog(this);
                File guarda =file.getSelectedFile();
                
                try {
                        FileOutputStream output = new FileOutputStream(guarda+".xls");
                        workbook.write(output);
                        output.close();
                } catch(Exception e) {
                        e.printStackTrace();
                }
                
        jTextArea1.setText("");
    }                                        

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {                                         
        // TODO add your handling code here:
        serialPort.removeEventListener();
	serialPort.close();
	jButton1.setEnabled(false);
        jButton2.setEnabled(false);
        jButton3.setEnabled(true);
        jComboBox1.setEnabled(false);
    }                                        

    private void jComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {                                           
        // TODO add your handling code here:
        Object selectedItem = jComboBox1.getSelectedItem();

        String com = selectedItem.toString();
        
        jButton1.setEnabled(false);
        jButton2.setEnabled(true);
        jButton3.setEnabled(false);
        jComboBox1.setEnabled(false);
        
        CommPortIdentifier portId = null;
        Enumeration portEnum = CommPortIdentifier.getPortIdentifiers();

		//First, Find an instance of serial port as set in PORT_NAMES.
		while (portEnum.hasMoreElements()) {
			CommPortIdentifier currPortId = (CommPortIdentifier) portEnum.nextElement();
				if (currPortId.getName().equals(com)) {
					portId = currPortId;
					break;
				}

		}
		if (portId == null) {
			System.out.println("Could not find COM port.");
			return;
		}

		try {
			// open serial port, and use class name for the appName.
			serialPort = (SerialPort) portId.open(this.getClass().getName(),
					TIME_OUT);

			// set port parameters
			serialPort.setSerialPortParams(DATA_RATE,
					SerialPort.DATABITS_8,
					SerialPort.STOPBITS_1,
					SerialPort.PARITY_NONE);

			// open the streams
			input = new BufferedReader(new InputStreamReader(serialPort.getInputStream()));
			//output = serialPort.getOutputStream();

			// add event listeners
			serialPort.addEventListener(this);
			serialPort.notifyOnDataAvailable(true);
		} catch (Exception e) {
			System.err.println(e.toString());
		}
    }                                          

    public synchronized void close() {
		if (serialPort != null) {
			serialPort.removeEventListener();
			serialPort.close();
		}
	}

	/**
	 * Handle an event on the serial port. Read the data and print it.
	 */
	public synchronized void serialEvent(SerialPortEvent oEvent) {
            Calendar calendario= Calendar.getInstance();
                
		if (oEvent.getEventType() == SerialPortEvent.DATA_AVAILABLE) {
			try {
				String inputLine=input.readLine();
                                temperatura[aux]= Double.parseDouble(inputLine);
                                hora[aux] =calendario.get(Calendar.HOUR_OF_DAY);
                                minutos [aux] = calendario.get(Calendar.MINUTE);
                                segundos [aux] = calendario.get(Calendar.SECOND);
                                jTextArea1.setText("\n\n        Hora: "+hora[aux]+":"+minutos[aux]+":"+segundos[aux]+"          Temperatura: "+temperatura[aux]+"°C");
                                if (temperatura[aux]<18)
                                {
                                    jTextArea1.setBackground(java.awt.Color.YELLOW);
                                    try
                                    {
                                        String sonido="C:/Users/PARCIF3/Documents/NetBeansProjects/Registrador_de_Temperatura 2/Ring01.wav";
                                        InputStream in = new FileInputStream (sonido);
                                        AudioStream audio;
                                        audio = new AudioStream (in);
                                        AudioPlayer.player.start(audio);

                                    }catch(Exception ex) {
                                        System.out.println (ex);
                                    }
                                }
                                else if (temperatura[aux]>28)
                                {
                                    jTextArea1.setBackground(java.awt.Color.ORANGE);
                                    try
                                    {
                                        String sonido="C:/Users/PARCIF3/Documents/NetBeansProjects/Registrador_de_Temperatura 2/Ring01.wav";
                                        InputStream in = new FileInputStream (sonido);
                                        AudioStream audio;
                                        audio = new AudioStream (in);
                                        AudioPlayer.player.start(audio);

                                    }catch(Exception ex) {
                                        System.out.println (ex);
                                    }
                                }
                                
                                else
                                    jTextArea1.setBackground(java.awt.Color.green);
                                aux=aux+1;
			} catch (Exception e) {
				System.err.println(e.toString());
			}
		}
		// Ignore all the other eventTypes, but you should consider the other ones.
	}

    public static void main(String args[]) throws Exception {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(Arduino.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Arduino.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Arduino.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Arduino.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        
        
        
        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Arduino().setVisible(true);
            }
        });
        
       //  AWTEvent a =java.awt.EventQueue.getCurrentEvent();
    }

    // Variables declaration - do not modify                     
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JComboBox jComboBox1;
    private javax.swing.JFileChooser jFileChooser1;
    private javax.swing.JFileChooser jFileChooser2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextArea jTextArea1;
    // End of variables declaration                   
}
