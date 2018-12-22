package mj.hotels.main;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ListSelectionModel;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatPart;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lib.FileDrop;

public class MainSwingController implements ActionListener {
	
	private JFrame frame = new JFrame();
	private JTable table;
    
    private String[] filePath = new String[15];
    
    //private Cell cell[][][] = new Cell[15][3000][2];
    private Cell cell[][] = new Cell[15][3000];
    private String fileInfo[][] = new String[15][2];
    
    private String overlapName[] = new String[3000];
    private int overlapCnt[] = new int[3000];
    
	
	protected String homeDir = System.getProperty("user.home") + "/Desktop";
	
    JButton process = new JButton("실행하기");
    JButton getFiles = new JButton("파일 불러오기");
    JButton makeXlsx = new JButton("엑셀 파일로 내보내기");
    
	int row, col;
    int fileCnt = 0;
    int pointer = 0;
    int overlapPointer = 0;
	int overlap = 0;
	// initialized
	
    public MainSwingController() {

        frame.setSize(new Dimension(720, 540));
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setTitle("호텔 중복 예약자 검색 프로그램 v1.1");
        frame.setLocationRelativeTo(null);

        initRightPane();
        initCenterPane();
        
        frame.setVisible(true);
    }

    private void initCenterPane() {
        JPanel centerPane = new JPanel();
        centerPane.setSize(740, frame.getHeight());

        String title[] = {"파일명", "파일 위치"};
        table = new JTable(fileInfo, title); 
        JScrollPane sp = new JScrollPane(table);
        centerPane.add(sp, BorderLayout.CENTER);

        // center align
        DefaultTableCellRenderer align = new DefaultTableCellRenderer();
        align.setHorizontalAlignment(SwingConstants.LEFT);
        frame.add(centerPane, BorderLayout.CENTER);

        // drag and drop 
        new FileDrop(System.out, centerPane, new FileDrop.Listener() {
            public void filesDropped(java.io.File[] files) {
                for (File file : files) {
                    String fileName = file.getName();
                    filePath[fileCnt] = file.getAbsolutePath();
                    fileInfo[fileCnt][0] = fileName;
                    fileInfo[fileCnt][1] = filePath[fileCnt];
                    
                    fileDroped(fileName, filePath[fileCnt]);
                    fileCnt++;
                    System.out.println("fileCnt : " + fileCnt + "\n");
                }
            }
        });

    }

    private void initRightPane() {
        JPanel rightPane = new JPanel();
        rightPane.setLayout(new GridLayout(11, 1));
        rightPane.setBounds(0, 0, 220, frame.getHeight());
        rightPane.setBackground(new Color(Integer.parseInt("B0BEC5", 16)));

        
        rightPane.add(getFiles);
        rightPane.add(process);
        rightPane.add(makeXlsx);
        
        process.addActionListener(this);
        getFiles.addActionListener(this);
        makeXlsx.addActionListener(this);
        
        
        frame.setLayout(new BorderLayout());
        frame.add(rightPane, BorderLayout.EAST);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
		if (e.getSource().equals(process)) {
			findNames();
			for (int i = 0; i < overlapPointer; i++)
				if (overlapCnt[i] > 1)
					overlap++;
	    	JOptionPane.showMessageDialog(null, overlap + "개의 중복이 발견되었습니다.");
		} else if (e.getSource().equals(getFiles)) {
			openFile();
		} else if (e.getSource().equals(makeXlsx)) {
			makeXlsx();
	    	JOptionPane.showMessageDialog(null, "바탕화면에 해당 파일이 생성되었습니다.");
		}
		
	}

    private void openFile() {
		JFileChooser chooser = new JFileChooser(homeDir);
		chooser.setDialogTitle("파일 열기");
		chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);
		FileNameExtensionFilter filter = new FileNameExtensionFilter("xlsx 문서 (2007년 이후 버전)" , "xlsx");
		chooser.setFileFilter(filter);	
		
		int isOpen = chooser.showOpenDialog(null);
		if (isOpen == JFileChooser.APPROVE_OPTION) {
			System.out.println("파일 불러오기 성공.");
			String fileName = chooser.getSelectedFile().getName();
			String filePath = chooser.getSelectedFile().getPath();
			fileInfo[fileCnt][0] = fileName;
            fileInfo[fileCnt][1] = filePath;
            
            fileDroped(fileName, filePath);
            fileCnt++;
            System.out.println("fileCnt : " + fileCnt + "\n");
		}
		table.repaint();
	}
	
    private void fileDroped(String name, String path) {
    	try {
            FileInputStream excelFile = new FileInputStream(new File(path));
            Workbook workbook = new XSSFWorkbook(excelFile);
            org.apache.poi.ss.usermodel.Sheet dataTypeSheet = workbook.getSheetAt(0);
            
            row = 0; 
            col = 0;
           
            int rows = 0;
            for (Row currentRow : dataTypeSheet) {
            	if (rows < 2) {
            		rows++;
            		continue;
            	}
                int tmp=0;
                
                boolean firstCol = true;
                for (Cell currentCell : currentRow) {
                	if (firstCol == true) {
                		firstCol = false;
                		continue;
                	}
                	
                	tmp++;
                	// print all of cells
                	//System.out.println(currentCell.getStringCellValue());
                	
                	cell[fileCnt][row] = currentCell;
                	
                	// read only one rows : Booker's name
                	if (tmp == 1)
                		break;
                }
                row++;
            }
         
            workbook.close();
            excelFile.close();
        } catch (FileNotFoundException e) {
        	JOptionPane.showMessageDialog(null, "파일을 찾을 수 없습니다.");
            e.printStackTrace();
        } catch (IOException e) {
        	JOptionPane.showMessageDialog(null, "오류가 발생하였습니다. 개발자에게 문의하세요.");
            e.printStackTrace();
        }
    }
		
    private void findNames() {
    	String name = null;
    	
    	for (pointer = 0; pointer < fileCnt; pointer++) {
    		
    		for (int i = 0; i < row; i++) {
    			if (cell[pointer][i] != null) {
    				name = cell[pointer][i].toString();
    				//System.out.println("name" + i + ": " + name);
    			}
    			else 
    				break;
    			
    			for (int j = 0; j < row; j++) {
        			if (cell[pointer][j] == null) {
        				break;
        			}

        			//System.out.println("name : " + name + "    overlapName[" + j + "] : " + overlapName[j]);
        			if (name.equals(cell[pointer][j].toString())) {
        				saveOverlap(name);
        				break;
        			}
        			
    			}
    		}
    		
    	}
    	
    }

	private void saveOverlap(String name) {
		boolean isOverlaped = false;
		int overlapTmp = 0;
		
		for (int i = 0; i < overlapPointer + 1; i++) {
			if (name.equals(overlapName[i])) {
				isOverlaped = true;
				overlapTmp = i;
				break;
			}
		}
		
		if (isOverlaped == true) {
			overlapCnt[overlapTmp]++;
			isOverlaped = false;
		} else {
			overlapName[overlapPointer] = name;
			overlapCnt[overlapPointer]++;
			overlapPointer++;
		}
		
		for (int i = 0; i < overlapPointer; i++) {
			System.out.println("overlapName" + i + ": " + overlapName[i]);
			System.out.println("overlapCnt" + i + ": " + overlapCnt[i]);
			
		}
		
	}
		
	private void makeXlsx() {
		try {
			Workbook outputBook = new XSSFWorkbook();
			org.apache.poi.ss.usermodel.Sheet outputSheet = outputBook.createSheet("호텔 중복 예약자");
			
			// sheet's style
			CellStyle titleStyle = outputBook.createCellStyle();
			CellStyle textStyle = outputBook.createCellStyle();

			Font titleFont = outputBook.createFont();
			Font textFont = outputBook.createFont();
			
			titleFont.setFontName("나눔고딕");
			titleFont.setBold(true);
			textFont.setFontName("나눔명조");
			
			titleStyle.setAlignment(HorizontalAlignment.CENTER);
			titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			titleStyle.setFont(titleFont);
			titleStyle.setBorderBottom(BorderStyle.THIN);
			titleStyle.setBorderTop(BorderStyle.THIN);
			titleStyle.setBorderRight(BorderStyle.THIN);
			titleStyle.setBorderLeft(BorderStyle.THIN);
			textStyle.setAlignment(HorizontalAlignment.CENTER);
			textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			textStyle.setFont(textFont);
			textStyle.setBorderBottom(BorderStyle.THIN);
			textStyle.setBorderTop(BorderStyle.THIN);
			textStyle.setBorderRight(BorderStyle.THIN);
			textStyle.setBorderLeft(BorderStyle.THIN);
			
			
			Row row = null;
			Cell cell = null;

			row = outputSheet.createRow(0);
			cell = row.createCell(0);
			cell.setCellValue("예약자명");
			cell.setCellStyle(titleStyle);
			cell = row.createCell(1);
			cell.setCellValue("예약 횟수");
			cell.setCellStyle(titleStyle);
			int POINTER = 0;
			for (int i = 0; i < overlapPointer; i++) {
				outputSheet.setColumnWidth((short) 0, (short)8000);
				//outputSheet.autoSizeColumn(POINTER);
				outputSheet.setColumnWidth(POINTER, (outputSheet.getColumnWidth(i)) + (short)1024);
				if (overlapCnt[i] > 1) {
					row = outputSheet.createRow(POINTER+1);
						
					cell = row.createCell(0);
					cell.setCellValue(overlapName[i]);
					cell.setCellStyle(textStyle);
					cell = row.createCell(1);
					cell.setCellValue(overlapCnt[i]);
					cell.setCellStyle(textStyle);
					POINTER++;
					//System.out.println("POINTER : " + POINTER);
				}
				
			}
			FileOutputStream fos = new FileOutputStream(homeDir + "/output.xlsx");
			outputBook.write(fos);
		} catch (FileNotFoundException e) {
        	JOptionPane.showMessageDialog(null, "파일 경로를 찾을 수 없습니다. 개발자에게 문의하세요.");
			e.printStackTrace();
		} catch (IOException e) { 
        	JOptionPane.showMessageDialog(null, "오류가 발생하였습니다. 개발자에게 문의하세요.");
			e.printStackTrace();
		}
		
	}
	
	// end of source
	
}


