

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Date;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

	static int countFiles = 0;

	public static void main(String[] args) throws Exception {
		List<Employee> list = new ArrayList<Employee>();

		search(list);

		Collections.sort(list);

		DateTimeFormatter FOMATTER = DateTimeFormatter.ofPattern("dd/MM/yyyy hh:mm");
		LocalDateTime localDateTime = LocalDateTime.now();

		String ldtString = FOMATTER.format(localDateTime);

		try {
			File file = new File("recordPocitadloVykonnosti.txt");
			int temp = 0;
			while (file.exists()) {
				temp++;
				file = new File("recordPocitadloVykonnosti" + String.valueOf(temp) + ".txt");
			}

			file.createNewFile();
			FileWriter csvWriter = new FileWriter(file);

			for (int i = 0; i < list.size(); i++) {

				csvWriter.append(list.get(i).getName());
				if ((list.get(i).getName()).length() < 8) {
					csvWriter.append("\t\t\t Stravený čas: " + Double.toString(list.get(i).getTime()));
				} else {
					csvWriter.append("\t\t Stravený čas: " + Double.toString(list.get(i).getTime()));
				}
				csvWriter.append("\t\t Hodnota: " + Double.toString(list.get(i).getValue()));
				csvWriter.append("\n");
				csvWriter.append("\n");

			}
			csvWriter.append("\n");
			csvWriter.append("Úspěšně načteno: " + countFiles + " souborů. \n");
			csvWriter.append("Datum a čas: " + ldtString + ". ");

			csvWriter.flush();
			csvWriter.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	public static List<Employee> read(FileInputStream fis, List<Employee> list) throws IOException {

		String tempName;
		String cellValue;
		double tempValue;
		double tempTime;
		boolean check = true;

		CellAddress cellAddress = null;
		Row row = null;
		Cell cell = null;
		Sheet sheet = null;
		final String sheetName = "Dokončenost";

		Workbook workbook = new XSSFWorkbook(fis);

		for (int k = 0; k < workbook.getNumberOfSheets(); k++) {

			if (sheetName.equals(workbook.getSheetName(k)))
				sheet = workbook.getSheetAt(k);

		}

		for (int i = 3; i < sheet.getLastRowNum() + 2; i++) {

			cellValue = "L" + String.valueOf(i);
			cellAddress = new CellAddress(cellValue);
			row = sheet.getRow(cellAddress.getRow());

			if (row == null)
				continue;

			cell = row.getCell(cellAddress.getColumn());

			if (cell == null)
				continue;

			if (cell.getCellType() == CellType.STRING) {
				if (cell.getStringCellValue() == "")
					continue;

				tempName = cell.getStringCellValue().trim();

			} else {
				continue;
			}

			cellValue = "H" + String.valueOf(i);
			cellAddress = new CellAddress(cellValue);
			row = sheet.getRow(cellAddress.getRow());
			cell = row.getCell(cellAddress.getColumn());

			if (cell == null) {
				tempValue = 0;
			} else if (cell.getCellType() == CellType.NUMERIC) {
				tempValue = cell.getNumericCellValue();
			} else if (cell.getCellType() == CellType.FORMULA) {

				tempValue = cell.getNumericCellValue();
			} else {
				tempValue = 0;
			}

			cellValue = "I" + String.valueOf(i);
			cellAddress = new CellAddress(cellValue);
			row = sheet.getRow(cellAddress.getRow());
			cell = row.getCell(cellAddress.getColumn());

			if (cell == null) {
				tempTime = 0;
			} else if (cell.getCellType() == CellType.NUMERIC) {
				tempTime = cell.getNumericCellValue();
			} else if (cell.getCellType() == CellType.FORMULA) {

				tempTime = cell.getNumericCellValue();
			} else {
				tempTime = 0;
			}

			for (int j = 0; j < list.size(); j++) {
				if (tempName.equals(list.get(j).getName())) {

					list.set(j, new Employee(tempName, (tempValue + list.get(j).getValue()),
							(tempTime + list.get(j).getTime())));

					check = false;
				}

			}
			if (check)
				list.add(new Employee(tempName, tempValue, tempTime));

			check = true;

		}

		workbook.close();
		fis.close();
		return (list);

	}

	public static List<Employee> search(List<Employee> list) throws IOException {

		final String directoryName = "C:\\Users\\vit.ropek\\Videos\\klienti";
		final String filePath = "dirIgnorePocitadlo.txt";
		String[] DirectoryIgnore;
		List<String> errorList = new ArrayList<String>();
		List<String> loadedList = new ArrayList<String>();
		boolean FEZExist = false;
		boolean file = false;

		if ((new File(filePath).exists())) {
			String content = null;
			try {
				content = Files.readString(Paths.get(filePath));
			} catch (IOException e) {
				e.printStackTrace();
			}
			DirectoryIgnore = content.split(System.getProperty("line.separator"));

		} else {
			DirectoryIgnore = new String[0];
		}

		File directory = new File(directoryName);
		String[] CompanyDirectory = directory.list();

		if (CompanyDirectory != null) {
			for (int i = 0; i < CompanyDirectory.length; i++) {
				String tempDir = directoryName + "\\" + CompanyDirectory[i];
				if (Arrays.asList(DirectoryIgnore).contains(tempDir))
					continue;

				File tempFile = new File(tempDir);

				String[] ProjectDirectory = tempFile.list();

				if (ProjectDirectory != null) {
					for (int j = 0; j < ProjectDirectory.length; j++) {
						String tempDirProject = tempDir + "\\" + ProjectDirectory[j];
						if (Arrays.asList(DirectoryIgnore).contains(tempDirProject))
							continue;

						File tempFileProject = new File(tempDirProject);
						String[] FileInProject = tempFileProject.list();

						if (FileInProject != null) {
							for (int k = 0; k < FileInProject.length; k++) {
								String filename = FileInProject[k];

								if (filename.equals("FEZ")) {
									String tempFEZ = tempDirProject + "\\" + FileInProject[k];
									FEZExist = true;
									if (Arrays.asList(DirectoryIgnore).contains(tempFEZ))
										continue;

									File tempDirFEZ = new File(tempFEZ);
									File[] files = tempDirFEZ.listFiles();
									Arrays.sort(files, Comparator.comparingLong(File::lastModified).reversed());

									for (int l = 0; l < files.length; l++) {
										String startName = "FEZ";
										String endName = "xlsx";
										String tempNameFile = files[l].getName();
										if(tempNameFile.length() > 7) {
										if (endName.equals(tempNameFile.substring(tempNameFile.length() - 4,
												tempNameFile.length()))) {
											if (startName.equals(tempNameFile.substring(0, 3))
													|| startName.equals(tempNameFile.substring(
															tempNameFile.length() - 8, tempNameFile.length() - 5))) {

												file = true;
												FileInputStream FEZFile = new FileInputStream(files[l]);
												String error;
												error = checkFile(FEZFile);
												FEZFile = new FileInputStream(files[l]);
												if (error.equals("null")) {
													read(FEZFile, list);
													countFiles++;
													loadedList.add(tempFEZ + "\\" + tempNameFile);
												} else {
													errorList.add(error);
													errorList.add(tempFEZ + "\\" + tempNameFile);

												}
												break;

											}
											}
										}
									}
									if (file == false) {
										errorList.add("Excelový soubor FEZ nenalezen");
										errorList.add(tempFEZ);
									}

								}
								file = false;
							}
							if (FEZExist == false) {
								errorList.add("Adresar FEZ nenalezen");
								errorList.add(tempDirProject);
							}
						}
						FEZExist = false;
					}
				}

			}
		}

		try {
			FileWriter csvWritere = new FileWriter("logPocitadloVykonnosti.txt");

			for (int m = 0; m < errorList.size(); m++) {

				csvWritere.append(errorList.get(m));
				csvWritere.append("\n");
				m++;
				csvWritere.append(errorList.get(m));
				csvWritere.append("\n\n");

			}

			csvWritere.append("Úspěšně načtené soubory: \n");
			for (int n = 0; n < loadedList.size(); n++) {

				csvWritere.append(loadedList.get(n));
				csvWritere.append("\n");

			}
			csvWritere.flush();
			csvWritere.close();

		} catch (IOException e) {

			e.printStackTrace();
		}

		return (list);

	}

	public static String checkFile(FileInputStream fis) throws IOException {

		String error;
		boolean data = false;
		CellAddress cellAddress = null;
		Row row = null;
		Cell cell = null;
		Sheet sheet = null;
		final String sheetName = "Dokončenost";
		final String checkL2 = "Responsible AMI";
		final String checkH2 = "Vytvořená hodnota (dle EVA) (MD)";
		final String checkI2 = "Strávený čas dle JIRA (MD)";
		
		try {
		Workbook work = new XSSFWorkbook(fis);

		for (int k = 0; k < work.getNumberOfSheets(); k++) {

			if (sheetName.equals(work.getSheetName(k)))
				sheet = work.getSheetAt(k);

		}
		if (sheet == null) {
			error = "List Dokončenost nenalezen. ";

			work.close();
			fis.close();
			return (error);
		}

		cellAddress = new CellAddress("L2");
		row = sheet.getRow(cellAddress.getRow());

		if (row == null) {
			error = "Sloupec Responsible AMI nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}
		cell = row.getCell(cellAddress.getColumn());

		if (cell == null) {
			error = "Sloupec Responsible AMI nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (cell.getCellType() != CellType.STRING) {
			error = "Sloupec Responsible AMI nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (checkL2.equals((cell.getStringCellValue()).trim()) == false) {
			error = "Sloupec Responsible AMI nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}
		cellAddress = new CellAddress("H2");
		row = sheet.getRow(cellAddress.getRow());
		cell = row.getCell(cellAddress.getColumn());

		if (cell == null) {
			error = "Sloupec Vytvořená hodnota (dle EVA) (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (cell.getCellType() != CellType.STRING) {
			error = "Sloupec Vytvořená hodnota (dle EVA) (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (checkH2.equals((cell.getStringCellValue()).trim()) == false) {
			error = "Sloupec Vytvořená hodnota (dle EVA) (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		cellAddress = new CellAddress("I2");
		row = sheet.getRow(cellAddress.getRow());
		cell = row.getCell(cellAddress.getColumn());

		if (cell == null) {
			error = "Sloupec Strávený čas dle JIRA (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (cell.getCellType() != CellType.STRING) {
			error = "Sloupec Strávený čas dle JIRA (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		if (checkI2.equals((cell.getStringCellValue()).trim()) == false) {
			error = "Sloupec Strávený čas dle JIRA (MD) nenalezen. ";
			work.close();
			fis.close();
			return (error);
		}

		for (int i = 3; i < sheet.getLastRowNum() + 2; i++) {

			String cellValue = "L" + String.valueOf(i);
			cellAddress = new CellAddress(cellValue);
			row = sheet.getRow(cellAddress.getRow());

			if (row == null)
				continue;

			cell = row.getCell(cellAddress.getColumn());

			if (cell == null)
				continue;

			if (cell.getCellType() == CellType.STRING) {
				if (cell.getStringCellValue() == "")
					continue;

				data = true;

			} else {
				continue;
			}

		}

		if (data == false) {
			error = "Na listu Dokončenost nenalezena žádná jména. ";
			work.close();
			fis.close();
			return (error);
		}

		error = "null";
		work.close();

		fis.close();
		return (error);
		}catch (Exception e) {
			error = "Soubor nelze otevřít. ";
			return(error);
		}
	}

}

class Employee implements Comparable {

	private String name;
	private double value;
	private double time;

	Employee(String name, double value, double time) {
		this.value = value;
		this.name = name;
		this.time = time;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public double getValue() {
		return value;
	}

	public void setValue(double value) {
		this.value = value;
	}

	public double getTime() {
		return time;
	}

	public void setTime(double time) {
		this.time = time;
	}

	@Override
	public int compareTo(Object o) {
		return this.getName().compareTo(((Employee) o).getName());
	}
}
