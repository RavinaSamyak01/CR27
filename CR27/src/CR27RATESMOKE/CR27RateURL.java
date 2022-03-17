package CR27RATESMOKE;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.annotations.Test;

import com.google.gson.*;

import BasePackage.BaseInit;
import BasePackage.Email;

public class CR27RateURL extends BaseInit {
	static JavascriptExecutor mzexecutor;

	public static int rcount;
	public static int f;
	public static int z;

	public static long totalPieces, pkgid, declaredValue, pkgweight, pkgheight, pkgwidth, pkglength;

	@Test
	public static void cr27RateURL() throws Exception {

		File srcCR = new File("./DataFile/CR27RATEJSON.xls");
		FileInputStream fisCR = new FileInputStream(srcCR);
		Workbook workbookCR = WorkbookFactory.create(fisCR);
		Sheet shCR = workbookCR.getSheet("Sheet1");
		rcount = shCR.getLastRowNum();

		for (f = 1; f <= rcount; f++) {

			File src = new File("./DataFile/CR27RATEJSON.xls");

			FileInputStream fis = new FileInputStream(src);
			Workbook workbook = WorkbookFactory.create(fis);
			Sheet sh1 = workbook.getSheet("Sheet1");

			DataFormatter formatter = new DataFormatter();

			/*
			 * 
			 * String vcd = "DEFG1234"; String uName = "ABCDE12345"; String pword =
			 * "yAVa9e7T9fb7hXVcMEqbJ9TFZ"; String ky = "xE9JlZWYbqXnJXjA"; String AccNumber
			 * = "528705287";
			 * 
			 * String pcomp = "CREATIVE ARTISTS AGENCY";
			 * 
			 * String padd1 = "CREATIVE ART CORP"; String padd2 = "55 STEVEN STREET"; String
			 * pcity = "GLADSTONE"; String pstate = "OR"; String pzip = "97034"; String
			 * pcountry = "USA"; String pres = "NO"; String popenhrs = "07:00"; String
			 * pclosehrs = "18:00"; String paddval = "YES";
			 * 
			 * String pcont = "John"; String pphone = "123-123-1233"; String pinst =
			 * "Contains glass items too.";
			 * 
			 * String dcomp = "CAMINO";
			 * 
			 * String dadd1 = "10 Summer St 6th Floor"; String dadd2 = "LAKE GROVE"; String
			 * dcity = "AURORA"; String dstate = "OR"; String dzip = "97027"; String
			 * dcountry = "USA"; String dres = "NO"; String dopenhrs = "07:00"; String
			 * dclosehrs = "18:00"; String daddval = "YES";
			 * 
			 * String dcont = "John"; String dphone = "123-123-1233"; String dinst =
			 * "Contains glass items too.";
			 * 
			 * String ServId = "PR";
			 * 
			 * int refId = 1; String refValue = "Ref1";
			 * 
			 * String shipDate = "05/19/2018"; String readyTime = "11:11"; int totalPieces =
			 * 1;
			 * 
			 * int pkgid = 1; int pkglength = 2; int pkgwidth = 3; int pkgheight = 4; int
			 * pkgweight = 5;
			 * 
			 * int declaredValue = 455; String promoCode = "";
			 * 
			 * String serviceFeatures = "SDRTS"; String signatureService = "SR";
			 * 
			 * String shpemail = "pdoshi@samyak.com"; String shpordrcv = "YES"; String shppu
			 * = "YES"; String shpdl = "YES"; String shpqdt = "YES"; String shpexcp = "YES";
			 * 
			 * String consemail = "pdoshi@samyak.com"; String consordrcv = "YES"; String
			 * conspu = "YES"; String consdl = "YES"; String consqdt = "YES"; String
			 * consexcp = "YES";
			 * 
			 * 
			 */

			String vcd = formatter.formatCellValue(sh1.getRow(f).getCell(0));
			String uName = formatter.formatCellValue(sh1.getRow(f).getCell(1));
			String pword = formatter.formatCellValue(sh1.getRow(f).getCell(2));
			String ky = formatter.formatCellValue(sh1.getRow(f).getCell(3));
			String AccNumber = formatter.formatCellValue(sh1.getRow(f).getCell(4));

			String pcomp = formatter.formatCellValue(sh1.getRow(f).getCell(5));
			String padd1 = formatter.formatCellValue(sh1.getRow(f).getCell(6));
			String padd2 = formatter.formatCellValue(sh1.getRow(f).getCell(7));
			String pcity = formatter.formatCellValue(sh1.getRow(f).getCell(8));
			String pstate = formatter.formatCellValue(sh1.getRow(f).getCell(9));
			String pzip = formatter.formatCellValue(sh1.getRow(f).getCell(10));
			String pcountry = formatter.formatCellValue(sh1.getRow(f).getCell(11));
			String pres = formatter.formatCellValue(sh1.getRow(f).getCell(12));
			String popenhrs = formatter.formatCellValue(sh1.getRow(f).getCell(13));
			String pclosehrs = formatter.formatCellValue(sh1.getRow(f).getCell(14));
			String paddval = formatter.formatCellValue(sh1.getRow(f).getCell(15));

			String dcomp = formatter.formatCellValue(sh1.getRow(f).getCell(16));
			String dadd1 = formatter.formatCellValue(sh1.getRow(f).getCell(17));
			String dadd2 = formatter.formatCellValue(sh1.getRow(f).getCell(18));
			String dcity = formatter.formatCellValue(sh1.getRow(f).getCell(19));
			String dstate = formatter.formatCellValue(sh1.getRow(f).getCell(20));
			String dzip = formatter.formatCellValue(sh1.getRow(f).getCell(21));
			String dcountry = formatter.formatCellValue(sh1.getRow(f).getCell(22));
			String dres = formatter.formatCellValue(sh1.getRow(f).getCell(23));
			String dopenhrs = formatter.formatCellValue(sh1.getRow(f).getCell(24));
			String dclosehrs = formatter.formatCellValue(sh1.getRow(f).getCell(25));
			String daddval = formatter.formatCellValue(sh1.getRow(f).getCell(26));

			String shipDate = formatter.formatCellValue(sh1.getRow(f).getCell(27));
			String readyTime = formatter.formatCellValue(sh1.getRow(f).getCell(28));
			String tpcs = formatter.formatCellValue(sh1.getRow(f).getCell(29));

			if (tpcs != "") {
				totalPieces = Long.parseLong(tpcs);
			} else {
				totalPieces = 0;
			}

			String pkgid1 = formatter.formatCellValue(sh1.getRow(f).getCell(30));
			if (pkgid1 != "") {
				pkgid = Long.parseLong(pkgid1);
			} else {
				pkgid = 0;
			}

			String pkglength1 = formatter.formatCellValue(sh1.getRow(f).getCell(31));
			if (pkglength1 != "") {
				pkglength = Long.parseLong(pkglength1);
			} else {
				pkglength = 0;
			}

			String pkgwidth1 = formatter.formatCellValue(sh1.getRow(f).getCell(32));
			if (pkgwidth1 != "") {
				pkgwidth = Long.parseLong(pkgwidth1);
			} else {
				pkgwidth = 0;
			}

			String pkgheight1 = formatter.formatCellValue(sh1.getRow(f).getCell(33));
			if (pkgheight1 != "") {
				pkgheight = Long.parseLong(pkgheight1);
			} else {
				pkgheight = 0;
			}

			String pkgweight1 = formatter.formatCellValue(sh1.getRow(f).getCell(34));
			if (pkgweight1 != "") {
				pkgweight = Long.parseLong(pkgweight1);
			} else {
				pkgweight = 0;
			}

			String dv = formatter.formatCellValue(sh1.getRow(f).getCell(35));
			if (dv != "") {
				declaredValue = Long.parseLong(dv);
			} else {
				declaredValue = 0;
			}

			String promoCode = formatter.formatCellValue(sh1.getRow(f).getCell(36));

			String serviceFeatures = formatter.formatCellValue(sh1.getRow(f).getCell(37));
			String signatureService = formatter.formatCellValue(sh1.getRow(f).getCell(38));

			Thread.sleep(1000);

			// Initial node
			// String InitialNode =
			// "http://10.20.205.70:9068/OrderShipRate/GetDetailRateRequest?format=json&json=";
			// DEV
			// String InitialNode =
			// "https://testws.fedexsameday.com/OrderShipRate/GetDetailRateRequest?format=json&json=";
			// Staging
			String InitialNode = "https://stagingws.fedexsameday.com/OrderShipRate/GetDetailRateRequest?format=json&json=";
			// TEMP Production
			// String InitialNode =
			// "https://webservicesda2.fedexsameday.com/OrderShipRate/GetDetailRateRequest?format=json&json=";
			// System.out.println(InitialNode);

			// User Credential Code
			JsonObject ucinfo = new JsonObject();
			ucinfo.addProperty("vcid", vcd);
			ucinfo.addProperty("userName", uName);
			ucinfo.addProperty("password", pword);
			ucinfo.addProperty("key", ky);
			ucinfo.addProperty("billingAccountNumber", AccNumber);

			JsonObject ucreden = new JsonObject();
			ucreden.add("userCredential", ucinfo);

			// 2 Portion - pickup details

			// B. Pickup Information
			JsonObject pinf = new JsonObject();
			pinf.addProperty("address1", padd1);
			pinf.addProperty("address2", padd2);
			pinf.addProperty("city", pcity);
			pinf.addProperty("state", pstate);
			pinf.addProperty("zipCode", pzip);
			pinf.addProperty("country", pcountry);
			pinf.addProperty("res", pres);
			pinf.addProperty("hrsOfOperationOpen", popenhrs);
			pinf.addProperty("hrsOfOperationClose", pclosehrs);
			pinf.addProperty("addressValidationOverride", paddval);

			JsonObject PUobj1 = new JsonObject();
			PUobj1.addProperty("pickupCompany", pcomp);
			PUobj1.add("pickupAddress", pinf);

			JsonObject PUobj2 = new JsonObject();
			PUobj2.add("pickup", PUobj1);

			// 3 Portion - delivery details

			// B. Delivery Information
			JsonObject dinf = new JsonObject();
			dinf.addProperty("address1", dadd1);
			dinf.addProperty("address2", dadd2);
			dinf.addProperty("city", dcity);
			dinf.addProperty("state", dstate);
			dinf.addProperty("zipCode", dzip);
			dinf.addProperty("country", dcountry);
			dinf.addProperty("res", dres);
			dinf.addProperty("hrsOfOperationOpen", dopenhrs);
			dinf.addProperty("hrsOfOperationClose", dclosehrs);
			dinf.addProperty("addressValidationOverride", daddval);

			JsonObject DLobj1 = new JsonObject();
			DLobj1.addProperty("deliveryCompany", dcomp);
			DLobj1.add("deliveryAddress", dinf);

			JsonObject DLobj2 = new JsonObject();
			DLobj2.add("delivery", DLobj1);

			// 4 Portion - Package info

			JsonObject pcslist = new JsonObject();
			pcslist.addProperty("id", pkgid);
			pcslist.addProperty("length", pkglength);
			pcslist.addProperty("width", pkgwidth);
			pcslist.addProperty("height", pkgheight);
			pcslist.addProperty("weight", pkgweight);

			JsonArray plst = new JsonArray();
			plst.add(pcslist);

			JsonObject ShpDtl = new JsonObject();
			ShpDtl.addProperty("shipDate", shipDate);
			ShpDtl.addProperty("readyTime", readyTime);
			ShpDtl.addProperty("totalPieces", totalPieces);

			if (totalPieces > 1) {
				JsonArray multiplst = new JsonArray();

				for (z = 1; totalPieces >= z; z++) {
					multiplst.add(pcslist);
				}

				ShpDtl.add("piecesList", multiplst);

			} else {
				ShpDtl.add("piecesList", plst);
			}

			ShpDtl.addProperty("declaredValue", declaredValue);
			ShpDtl.addProperty("promoCode", promoCode);

			JsonObject ShipDetails = new JsonObject();
			ShipDetails.add("shipmentDetail", ShpDtl);

			// Service Feature

			JsonArray sftrlist = new JsonArray();
			sftrlist.add(serviceFeatures);

			JsonObject SF = new JsonObject();
			SF.add("serviceFeatures", sftrlist);

			// Signature Service

			JsonObject SS = new JsonObject();
			SS.addProperty("signatureService", signatureService);

			// Final Object

			Object device = new Object();
			JsonObject deviceJson = new Gson().toJsonTree(device).getAsJsonObject();

			for (Map.Entry<String, JsonElement> entry : SS.entrySet()) {
				for (Map.Entry<String, JsonElement> entry1 : SF.entrySet()) {
					for (Map.Entry<String, JsonElement> entry2 : ShipDetails.entrySet()) {
						for (Map.Entry<String, JsonElement> entry3 : DLobj2.entrySet()) {
							for (Map.Entry<String, JsonElement> entry4 : PUobj2.entrySet()) {
								for (Map.Entry<String, JsonElement> entry5 : ucreden.entrySet()) {
									deviceJson.add(entry5.getKey(), entry5.getValue());
								}
								deviceJson.add(entry4.getKey(), entry4.getValue());
							}
							deviceJson.add(entry3.getKey(), entry3.getValue());
						}
						deviceJson.add(entry2.getKey(), entry2.getValue());
					}
					deviceJson.add(entry1.getKey(), entry1.getValue());
				}
				deviceJson.add(entry.getKey(), entry.getValue());
			}

			// System.out.println(deviceJson);

			String CR27RateURL = InitialNode + deviceJson;
			System.out.println(CR27RateURL);

			mzexecutor = (JavascriptExecutor) ChDriver;
			ChDriver.get(CR27RateURL);
			Thread.sleep(5000);
			getscreenshot();

			String AB = ChDriver.findElement(By.xpath("/html/body/pre")).getText();

			String FinalResult = AB;
			System.out.println("FinalResult : " + FinalResult);

			File src1 = new File("./DataFile/CR27RATEJSON.xls");
			FileOutputStream fis1 = new FileOutputStream(src1);
			Sheet sh2 = workbook.getSheet("Sheet1");
			sh2.getRow(f).createCell(40).setCellValue(AB);
			workbook.write(fis1);
			fis1.close();

			// Check Expected and Act Result
			String ExpResult = formatter.formatCellValue(sh1.getRow(f).getCell(39));

			// SELECT * FROM WebServiceErrorCodeMessage

			if (!AB.contains("LM")) {
				File src1srcstatus = new File("./DataFile/CR27RATEJSON.xls");
				FileOutputStream fis2 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(41).setCellValue("LM not valid");
				workbook.write(fis2);

				// close the file
				fis2.close();
			} else {
				File src1srcstatus = new File("./DataFile/CR27RATEJSON.xls");
				FileOutputStream fis3 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(41).setCellValue("LM is Valid");
				workbook.write(fis3);

				// close the file
				fis3.close();
			}

			if (!FinalResult.contains(ExpResult)) {
				File src1srcstatus = new File("./DataFile/CR27RATEJSON.xls");
				FileOutputStream fis2 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(42).setCellValue("FAIL");
				workbook.write(fis2);

				// close the file
				fis2.close();
			} else {
				File src1srcstatus = new File("./DataFile/CR27RATEJSON.xls");
				FileOutputStream fis3 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(42).setCellValue("PASS");
				workbook.write(fis3);

				// close the file
				fis3.close();
			}
		}

		msg.append("\n" + "Please Find attached sheet for CR27 Rate URL Result" + "\n\n");
		msg.append("*** This is automated generated email and send through automation script" + "\n");

		String subject = "Selenium Automation Script: CR27 Rate URL Process";
		try {
			Email.sendMail("ravina.prajapati@samyak.com,asharma@samyak.com,parth.doshi@samyak.com", subject, msg.toString(), "./DataFile/CR27RATEJSON.xls");
		} catch (Exception ex) {
			Logger.getLogger(CR27RateURL.class.getName()).log(Level.SEVERE, null, ex);
		}
		System.out.println("The End !!!!!!!!!!!!");
	}

	public static void getscreenshot() throws Exception {
		File scrFile = ((TakesScreenshot) ChDriver).getScreenshotAs(OutputType.FILE);
		// The below method will save the screen shot in d drive with name
		// "screenshot.png"

		int s = f;

		FileUtils.copyFile(scrFile, new File("./Rate_Screenshot/" + "Result" + s + ".png"));
	}
}