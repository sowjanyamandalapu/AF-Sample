package CanparPODDownload;


import pivac.utilities.Globals;
import pivac.utilities.StringPadding;

import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;
import java.util.stream.Stream;

import javax.imageio.ImageIO;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import net.sourceforge.tess4j.ITesseract;
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import pivac.messaging.EmailMessage;


public class App 
{
	static boolean isFirstClick = false;
	static String sourceFileName;
	static String sourceFileFolderName;
	public static void main(String[] args) throws Throwable {

		//WebDriverManager.chromedriver().setup();
		Globals.setAppName("Day&Ross 214 Automation");

		//System.setProperty("webdriver.chrome.driver", "C:\\Users\\ediadmin\\chromedriver-win64\\chromedriver.exe");
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\smandalapu\\chromedriver-win64\\chromedriver.exe");

		ChromeDriver driver = new ChromeDriver(); //ChromeDriver driver

		driver.manage().window().maximize();
		driver.manage().deleteAllCookies();

		//send URL

		driver.manage().timeouts().implicitlyWait(Duration.ofMillis(3000));
		driver.get("https://dayross.com/en/track-shipments");
		driver.manage().timeouts().pageLoadTimeout(Duration.ofMillis(2000));
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofMillis(10000));


		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='sc-jlVhvo eTntCi']//button[@class='sc-brSamD coqlhD sc-cxYBSk eHburb']")));

		WebElement alert = driver.findElement(By.xpath("//div[@class='sc-jlVhvo eTntCi']//button[@class='sc-brSamD coqlhD sc-cxYBSk eHburb']"));
		alert.click();

		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='rw-input rw-dropdown-list-input']")));
		WebElement dropDown = driver.findElement(By.xpath("//div[@class='rw-input rw-dropdown-list-input']"));
		dropDown.click();


		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='rw-popup']//ul/li[3]")));
		WebElement selectLT = driver.findElement(By.xpath("//div[@class='rw-popup']//ul/li[3]"));
		selectLT.click();


		List<SendDetails> sendDetails = new ArrayList<>();

		try (Stream<Path> path = Files.list(Paths.get("C:\\Users\\smandalapu\\OneDrive - Pival International Inc\\Desktop\\Toyo\\Canpar\\Input"))) {


			path.filter(file -> !Files.isDirectory(file))
			.forEach(p -> {	
				Globals.debugMessage("sucessfully Read the all the files from the folder.");

				sourceFileFolderName = p.getParent().toString() + "\\" ;
				sourceFileName =  p.getFileName().toString();

				try (Stream<String> lines = Files.lines(p, StandardCharsets.UTF_8)){
					Globals.debugMessage("Successfully read the data file ");

					lines.skip(1)
					.filter(f -> !f.startsWith("Load"))
					.forEach(f -> {	
						Globals.debugMessage("Successfully enter in lambdas expression ");


						try {
							Globals.debugMessage("try block ");

							String[] fields = f.split(",");	

							String proNumber = fields[1];

							System.out.println("-----------------------------");
							System.out.println("proNumber:"+proNumber);

							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[normalize-space()='Reset']")));
							WebElement reset =  driver.findElement(By.xpath("//button[normalize-space()='Reset']"));
							reset.click();

							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//textarea[@name='shipmentNumbers']")));

							WebElement proNumberField = driver.findElement(By.xpath("//textarea[@name='shipmentNumbers']"));
							proNumberField.clear();

							proNumberField.sendKeys(proNumber);

							String orginalWindow = driver.getWindowHandle();

							wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[@type='submit']")));

							WebElement search =  driver.findElement(By.xpath("//button[@type='submit']"));
							search.click();

							Thread.sleep(3000);

							for(String windowHandle : driver.getWindowHandles()) {
								if(!windowHandle.endsWith(orginalWindow)) {
									driver.switchTo().window(windowHandle);
									driver.close();
									driver.switchTo().window(orginalWindow);
									System.out.println("Orginal Handler:"+orginalWindow);
								}
							}

							try {
								WebElement status = driver.findElement(By.xpath("//td[5]"));
								String statusMessage =  status.getText();

								System.out.println("Order Status:"+statusMessage);
								if(statusMessage.equalsIgnoreCase("Delivered")) {
									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[4]")));
									String date =  driver.findElement(By.xpath("//td[4]")).getText();
									
									DateTimeFormatter inputformatter = DateTimeFormatter.ofPattern("MM-dd-yyyy (HH:mm)");
									DateTimeFormatter outputformatter = DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss");
									LocalDateTime dateTime = LocalDateTime.parse(date,inputformatter);
									String formatedDate = dateTime.format(outputformatter);
									
									System.out.println("Order Date:"+formatedDate);

									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//img[@alt='expand filter']")));
									WebElement expand =  driver.findElement(By.xpath("//img[@alt='expand filter']"));
									expand.click();


									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='column left']//span[@class='city']")));
									String fromCity =  driver.findElement(By.xpath("//div[@class='column left']//span[@class='city']")).getText();
									System.out.println("Order fromcity:"+fromCity);

									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='column right']//span[@class='postalcode']")));
									String fromPostalcode =  driver.findElement(By.xpath("//div[@class='column left']//span[@class='postalcode']")).getText();
									System.out.println("Order fromPostalcode:"+fromPostalcode);


									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='column left']//span[@class='city']")));
									String toCity =  driver.findElement(By.xpath("//div[@class='column left']//span[@class='city']")).getText();
									System.out.println("Order toCity:"+toCity);

									wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[@class='column right']//span[@class='city']")));
									String toPostalcode =  driver.findElement(By.xpath("//div[@class='column right']//span[@class='postalcode']")).getText();
									System.out.println("Order toPostalcode:"+toPostalcode);
									
									Thread.sleep(2000);
									
									Globals.debugMessage("Writting record out to the destination lication.");
									String filename = UUID.randomUUID().toString().toUpperCase();
									try (FileWriter writer = new FileWriter("C:\\Users\\smandalapu\\OneDrive - Pival International Inc\\Desktop\\Toyo\\Canpar\\Output\\"+ proNumber + "_" + filename + ".txt");
										BufferedWriter bWriter = new BufferedWriter(writer)) {
										String finalData = proNumber + "|" + fields[0] + "|" + "|" + formatedDate
										+ "|" + "D1" + "|" + "NS" + "|" + fromCity.toUpperCase() + "|" + fromPostalcode.split(" ")[0].toUpperCase() 
										+ "|" + toCity.toUpperCase() + "|" + toPostalcode.split(" ")[0].toUpperCase() + "|" + "|" + "|";
										
										bWriter.write(finalData);	
										Thread.sleep(2000);				
									}

								}else {
									System.out.println("Order status:"+statusMessage);
									sendDetails.add(new SendDetails(proNumber, fields[0], statusMessage));

								}
							} catch (NoSuchElementException e) {

								WebElement error = driver.findElement(By.xpath("//h2[@class='sc-guDLey sc-hLQSwg sc-hFtlrn cVtAjV hLhhMj diZljP']"));
								if(error.isDisplayed()) {
									String errorMessage = driver.findElement(By.xpath("//p[@class='sc-sANrS TKFCC']")).getText();
									System.out.println("Error message:"+errorMessage);
									sendDetails.add(new SendDetails(proNumber, fields[0], errorMessage));
								}

							}

						} catch (Exception e) {
							e.printStackTrace();
							EmailMessage.sendException(e);
							try {
								sendEmailNotification(sendDetails, driver);
							} catch (Exception e1) {
								e1.printStackTrace();
								EmailMessage.sendException(e);
								System.exit(128);
							}
							System.exit(128);
						}
					});
				} catch (Exception e) {
					e.printStackTrace();
					EmailMessage.sendException(e);
					try {
						sendEmailNotification(sendDetails, driver);
					} catch (Exception e1) {
						e1.printStackTrace();
						EmailMessage.sendException(e);
						System.exit(128);
					}
					System.exit(128);
				}
			});

		} catch (IOException e) {
			e.printStackTrace();
			EmailMessage.sendException(e);
			sendEmailNotification(sendDetails, driver);
			System.exit(128);
		} finally {
			sendEmailNotification(sendDetails, driver);
		}

	}

	private static void fileRename(final String newFileName, final String folder) throws IOException, InterruptedException {
		File folderName = new File(folder);
		if (folderName.isDirectory()) {
			File[] files = folderName.listFiles();

			for (File file : files) {
				if(file.isFile()) {
					if(file.getName().length() > 10 && !(file.getName().startsWith("9") || file.getName().startsWith("8"))) {
						//String name[] = file.getName().split("[.]");
						String newName = folder + "\\" + newFileName + ".pdf";
						file.renameTo(new File(newName));
					}
				}
			}
		}

	}
	private static void sendEmailNotification(List<SendDetails> sendDetails, WebDriver driver) throws Exception {

		driver.close();
		driver.quit();	

		//Dropping file into Archive

		if(sourceFileName != null) {
			String name[] = sourceFileName.split("[.]");

			//Dropping file into Archive
			DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
			LocalDate currentDate = LocalDate.now();
			String formattedDate = currentDate.format(dateTimeFormatter);

			String renamedFileName =  name[0] + "_" + formattedDate + "." + name[1];
			Path temp = Files.move
					(Paths.get(sourceFileFolderName + sourceFileName), 
							Paths.get(sourceFileFolderName.replaceAll("Input", "Archive")  + renamedFileName), StandardCopyOption.REPLACE_EXISTING);

			if(temp != null)
				Globals.debugMessage("File renamed and moved successfully");
		}

		//Sending email Notification

		if (sendDetails.size() >= 1) {

			Workbook scrExcel = new XSSFWorkbook();
			Sheet sheet = scrExcel.createSheet("Day&Ross 214 Automation Failure Report");

			CellStyle border = scrExcel.createCellStyle();
			border.setBorderBottom(BorderStyle.THIN);
			border.setBottomBorderColor(IndexedColors.BLACK.getIndex());
			border.setBorderRight(BorderStyle.THIN);
			border.setRightBorderColor(IndexedColors.BLACK.getIndex());
			border.setBorderLeft(BorderStyle.THIN);
			border.setLeftBorderColor(IndexedColors.BLACK.getIndex());
			border.setBorderTop(BorderStyle.THIN);
			border.setTopBorderColor(IndexedColors.BLACK.getIndex());

			Row row = sheet.createRow(0);

			int cellCounter = 0;
			Cell cell = row.createCell(cellCounter);
			cell.setCellStyle(border);
			cell.setCellValue("PRO #");

			cell = row.createCell(++cellCounter);
			cell.setCellStyle(border);
			cell.setCellValue("Error Message/Status");

			cell = row.createCell(++cellCounter);
			cell.setCellStyle(border);
			cell.setCellValue("Load #");

			int rowCounter = 0;
			for (SendDetails detail : sendDetails) {
				rowCounter++;
				row = sheet.createRow(rowCounter);

				cellCounter = 0;
				cell = row.createCell(cellCounter);
				cell.setCellValue(detail.getProNumber());

				cell = row.createCell(++cellCounter);
				cell.setCellValue(detail.getErrorMessage());

				cell = row.createCell(++cellCounter);
				cell.setCellValue(detail.getLoadNumber());
			}

			//auto-set all the column widths
			for (int i = 0; i < cellCounter; i++) {
				sheet.autoSizeColumn(i);
				sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 11/10);
			}

			String fileName = "C:\\Users\\smandalapu\\OneDrive - Pival International Inc\\Desktop\\Toyo\\Canpar\\" + "Day&Ross_214_Automation_Failure_Report.xlsx";

			//print and close the workbook
			try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
				scrExcel.write(fileOut);
			} finally {
				scrExcel.close();
			}

			//Sending email Notification

			EmailMessage message = new EmailMessage();
			LocalDate date = LocalDate.now();

			Globals.debugMessage("Started building email body");

			message.addRecipient("Edi@pival.com");
			message.addCcRecipient("Edi@pival.com");
			message.setSubject("Day&Ross 214 Automation Failure Report"); 

			//build the html body
			StringBuilder builder = new StringBuilder();
			builder.append("Please see the attached Day&Ross 214 Automation Failure report for " + date + ".<br><br>");
			builder.append("<font face=\"Courier New\"><b>");
			builder.append(StringPadding.padRight("", "&nbsp;", 12));
			builder.append("<br>");

			builder.append("<br></font>");

			message.addHtmlBody(builder.toString());
			message.setSender("pivalit@pival.com");
			message.setSendDateTime(LocalDateTime.now());
			message.addAttachment(new File(fileName));

			Globals.debugMessage("Sending email successfully....");

			message.postMessage();
		}

	}


}

/* ---------- Not in Usage----------
 * 
 * public static void convertPDFToTiff(String folderPath) throws IOException {
		String newfolder = "D:\\Java\\File Transfers\\ToECS\\VitranPOD\\Final";
		File folderName = new File(folderPath);
		if (folderName.isDirectory()) {
			File[] files = folderName.listFiles();

			for (File file : files) {
				if(file.isFile()) {

					String name[] = file.getName().split("[.]");
					String tiffPath = newfolder + "\\" + name[0] + ".TIF";
					String pdfPath = folderPath + "\\" + file.getName();
					// Load the PDF document
					PDDocument document = PDDocument.load(new File(pdfPath));
					// Create PDFRenderer to render PDF pages to images
					PDFRenderer pdfRenderer = new PDFRenderer(document);

					//Consignee
					for (int page = 0; page < document.getNumberOfPages(); ++page) {
						// Render the first page (you can iterate over all pages if needed)
						BufferedImage bufferedImage = pdfRenderer.renderImageWithDPI(page, 300); // Render at 300 DPI

						ITesseract instance = new Tesseract();
						instance.setLanguage("eng");
						String result;
						try {
							result = instance.doOCR(bufferedImage);

							if(result.contains("Consignee") || result.contains("CONSIGNEE") || result.contains("Proof Of Delivery") || result.contains("Delivered")) {
								// Write the buffered image to a TIFF file
								File outputfile = new File(tiffPath);
								ImageIO.write(bufferedImage, "TIFF", outputfile);
								break;

							}else {
							}
						} catch (TesseractException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
							System.out.println("Exception:" + e.getMessage());
						}



					}
					// Close the PDF document

					document.close();
					//file.delete();
				}
			}
		}
	}*/



