package CR27SMOKE;

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

public class CR27ShipURL extends BaseInit {

	static JavascriptExecutor mzexecutor;

	public static int rcount;
	public static int f;
	public static int z;

	public static long refId, totalPieces, pkgid, declaredValue, pkgweight, pkgheight, pkgwidth, pkglength;

	@Test
	public static void cr27ShipURL() throws Exception {

		File srcCR = new File("./DataFile/CR27SHIPJSON.xls");

		FileInputStream fisCR = new FileInputStream(srcCR);
		Workbook workbookCR = WorkbookFactory.create(fisCR);
		Sheet shCR = workbookCR.getSheet("Sheet1");
		rcount = shCR.getLastRowNum();

		for (f = 1; f <= rcount; f++) {

			File src = new File("./DataFile/CR27SHIPJSON.xls");

			FileInputStream fis = new FileInputStream(src);
			Workbook workbook = WorkbookFactory.create(fis);
			Sheet sh1 = workbook.getSheet("Sheet1");

			DataFormatter formatter = new DataFormatter();

			/*
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
			String pcont = formatter.formatCellValue(sh1.getRow(f).getCell(16));
			String pphone = formatter.formatCellValue(sh1.getRow(f).getCell(17));
			String pinst = formatter.formatCellValue(sh1.getRow(f).getCell(18));

			String dcomp = formatter.formatCellValue(sh1.getRow(f).getCell(19));
			String dadd1 = formatter.formatCellValue(sh1.getRow(f).getCell(20));
			String dadd2 = formatter.formatCellValue(sh1.getRow(f).getCell(21));
			String dcity = formatter.formatCellValue(sh1.getRow(f).getCell(22));
			String dstate = formatter.formatCellValue(sh1.getRow(f).getCell(23));
			String dzip = formatter.formatCellValue(sh1.getRow(f).getCell(24));
			String dcountry = formatter.formatCellValue(sh1.getRow(f).getCell(25));
			String dres = formatter.formatCellValue(sh1.getRow(f).getCell(26));
			String dopenhrs = formatter.formatCellValue(sh1.getRow(f).getCell(27));
			String dclosehrs = formatter.formatCellValue(sh1.getRow(f).getCell(28));
			String daddval = formatter.formatCellValue(sh1.getRow(f).getCell(29));
			String dcont = formatter.formatCellValue(sh1.getRow(f).getCell(30));
			String dphone = formatter.formatCellValue(sh1.getRow(f).getCell(31));
			String dinst = formatter.formatCellValue(sh1.getRow(f).getCell(32));

			String ServId = formatter.formatCellValue(sh1.getRow(f).getCell(33));

			String refId1 = formatter.formatCellValue(sh1.getRow(f).getCell(34));

			if (refId1 != "") {
				refId = Long.parseLong(refId1);
			} else {
				refId = 0;
			}

			String refValue = formatter.formatCellValue(sh1.getRow(f).getCell(35));

			String shipDate = formatter.formatCellValue(sh1.getRow(f).getCell(36));
			String readyTime = formatter.formatCellValue(sh1.getRow(f).getCell(37));
			String tpcs = formatter.formatCellValue(sh1.getRow(f).getCell(38));

			if (tpcs != "") {
				totalPieces = Long.parseLong(tpcs);
			} else {
				totalPieces = 0;
			}

			String pkgid1 = formatter.formatCellValue(sh1.getRow(f).getCell(39));

			if (pkgid1 != "") {
				pkgid = Long.parseLong(pkgid1);
			} else {
				pkgid = 0;
			}

			String pkglength1 = formatter.formatCellValue(sh1.getRow(f).getCell(40));

			if (pkglength1 != "") {
				pkglength = Long.parseLong(pkglength1);
			} else {
				pkglength = 0;
			}

			String pkgwidth1 = formatter.formatCellValue(sh1.getRow(f).getCell(41));

			if (pkgwidth1 != "") {
				pkgwidth = Long.parseLong(pkgwidth1);
			} else {
				pkgwidth = 0;
			}

			String pkgheight1 = formatter.formatCellValue(sh1.getRow(f).getCell(42));

			if (pkgheight1 != "") {
				pkgheight = Long.parseLong(pkgheight1);
			} else {
				pkgheight = 0;
			}

			String pkgweight1 = formatter.formatCellValue(sh1.getRow(f).getCell(43));

			if (pkgweight1 != "") {
				pkgweight = Long.parseLong(pkgweight1);
			} else {
				pkgweight = 0;
			}

			String dv = formatter.formatCellValue(sh1.getRow(f).getCell(44));

			if (dv != "") {
				declaredValue = Long.parseLong(dv);
			} else {
				declaredValue = 0;
			}

			String promoCode = formatter.formatCellValue(sh1.getRow(f).getCell(45));

			String serviceFeatures = formatter.formatCellValue(sh1.getRow(f).getCell(46));
			String signatureService = formatter.formatCellValue(sh1.getRow(f).getCell(47));

			String shpemail = formatter.formatCellValue(sh1.getRow(f).getCell(48));
			String shpordrcv = formatter.formatCellValue(sh1.getRow(f).getCell(49));
			String shppu = formatter.formatCellValue(sh1.getRow(f).getCell(50));
			String shpdl = formatter.formatCellValue(sh1.getRow(f).getCell(51));
			String shpqdt = formatter.formatCellValue(sh1.getRow(f).getCell(52));
			String shpexcp = formatter.formatCellValue(sh1.getRow(f).getCell(53));

			String consemail = formatter.formatCellValue(sh1.getRow(f).getCell(54));
			String consordrcv = formatter.formatCellValue(sh1.getRow(f).getCell(55));
			String conspu = formatter.formatCellValue(sh1.getRow(f).getCell(56));
			String consdl = formatter.formatCellValue(sh1.getRow(f).getCell(57));
			String consqdt = formatter.formatCellValue(sh1.getRow(f).getCell(58));
			String consexcp = formatter.formatCellValue(sh1.getRow(f).getCell(59));
			Thread.sleep(1000);

			// Initial node
			// String InitialNode =
			// "http://10.20.205.70:9068/OrderShipment/OrderShipRequest?format=json&json=";
			// DEV
			// String InitialNode =
			// "https://testws.fedexsameday.com/OrderShipment/OrderShipRequest?format=json&json=";
			// Staging
			String InitialNode = "https://stagingws.fedexsameday.com/OrderShipment/OrderShipRequest?format=json&json=";
			// TEMP Production
			// String InitialNode =
			// "https://webservicesda2.fedexsameday.com/OrderShipment/OrderShipRequest?format=json&json=";
			// System.out.println(InitialNode);

			// User Credential Code

			JsonObject ucinfo = new JsonObject();
			ucinfo.addProperty("vcid", vcd);
			// ucinfo.addProperty("vcid", sh1.getRow(f).getCell(0).getStringCellValue());
			ucinfo.addProperty("userName", uName);
			ucinfo.addProperty("password", pword);
			ucinfo.addProperty("key", ky);
			ucinfo.addProperty("billingAccountNumber", AccNumber);

			JsonObject ucreden = new JsonObject();
			ucreden.add("userCredential", ucinfo);

			// System.out.println(InitialNode);
			// System.out.println("User Credential : " + ucreden);

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
			PUobj1.addProperty("pickupContact", pcont);
			PUobj1.addProperty("pickupPhone", pphone);
			PUobj1.addProperty("pickupInstructions", pinst);

			JsonObject PUobj2 = new JsonObject();
			PUobj2.add("pickup", PUobj1);

			// System.out.println("PUobj2 : " + PUobj2);

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
			DLobj1.addProperty("deliveryContact", dcont);
			DLobj1.addProperty("deliveryPhone", dphone);
			DLobj1.addProperty("deliveryInstructions", dinst);

			JsonObject DLobj2 = new JsonObject();
			DLobj2.add("delivery", DLobj1);

			// System.out.println("DLobj2 : " + DLobj2);

			// 4 Portion - Service

			JsonObject srv = new JsonObject();
			srv.addProperty("serviceId", ServId);

			// System.out.println("serviceId : " + srv);

			// 5 Portion - Package info

			JsonObject rf = new JsonObject();
			rf.addProperty("refId", refId);
			rf.addProperty("refValue", refValue);

			JsonArray refList = new JsonArray();
			refList.add(rf);

			JsonObject pcslist = new JsonObject();
			pcslist.addProperty("id", pkgid);
			pcslist.addProperty("length", pkglength);
			pcslist.addProperty("width", pkgwidth);
			pcslist.addProperty("height", pkgheight);
			pcslist.addProperty("weight", pkgweight);

			JsonArray plst = new JsonArray();
			plst.add(pcslist);

			JsonObject ShpDtl = new JsonObject();
			ShpDtl.add("refList", refList);
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

			// ShpDtl.add("piecesList", plst);
			ShpDtl.addProperty("declaredValue", declaredValue);
			ShpDtl.addProperty("promoCode", promoCode);

			JsonObject ShipDetails = new JsonObject();
			ShipDetails.add("shipmentDetail", ShpDtl);

			// System.out.println("shipmentDetail : " + ShipDetails);

			// 6. Service Feature

			JsonArray sftrlist = new JsonArray();
			sftrlist.add(serviceFeatures);

			JsonObject SF = new JsonObject();
			SF.add("serviceFeatures", sftrlist);

			// System.out.println("serviceFeatures : " + SF);

			// 7. Signature Service

			JsonObject SS = new JsonObject();
			SS.addProperty("signatureService", signatureService);

			// System.out.println("signatureService : " + SS + "\n\n");

			// 8. shipperEmailNotification

			JsonObject shpemlnot = new JsonObject();
			shpemlnot.addProperty("emailAddress", shpemail);
			shpemlnot.addProperty("orderReceived", shpordrcv);
			shpemlnot.addProperty("pickup", shppu);
			shpemlnot.addProperty("delivered", shpdl);
			shpemlnot.addProperty("qdtChange", shpqdt);
			shpemlnot.addProperty("exception", shpexcp);

			JsonObject ShipEmailNotf = new JsonObject();
			ShipEmailNotf.add("shipperEmailNotification", shpemlnot);

			// System.out.println("shipperEmailNotification : " + ShipEmailNotf + "\n");

			// 9. consigneeEmailNotification

			JsonObject Consemlnot = new JsonObject();
			Consemlnot.addProperty("emailAddress", consemail);
			Consemlnot.addProperty("orderReceived", consordrcv);
			Consemlnot.addProperty("pickup", conspu);
			Consemlnot.addProperty("delivered", consdl);
			Consemlnot.addProperty("qdtChange", consqdt);
			Consemlnot.addProperty("exception", consexcp);

			JsonObject ConsEmailNotf = new JsonObject();
			ConsEmailNotf.add("consigneeEmailNotification", Consemlnot);

			// System.out.println("consigneeEmailNotification : " + ConsEmailNotf + "\n");

			// Final Object

			Object device = new Object();
			JsonObject deviceJson = new Gson().toJsonTree(device).getAsJsonObject();

			for (Map.Entry<String, JsonElement> entry : ConsEmailNotf.entrySet()) {
				for (Map.Entry<String, JsonElement> entry1 : ShipEmailNotf.entrySet()) {
					for (Map.Entry<String, JsonElement> entry2 : SS.entrySet()) {
						for (Map.Entry<String, JsonElement> entry3 : SF.entrySet()) {
							for (Map.Entry<String, JsonElement> entry4 : ShipDetails.entrySet()) {
								for (Map.Entry<String, JsonElement> entry5 : srv.entrySet()) {
									for (Map.Entry<String, JsonElement> entry6 : DLobj2.entrySet()) {
										for (Map.Entry<String, JsonElement> entry7 : PUobj2.entrySet()) {
											for (Map.Entry<String, JsonElement> entry8 : ucreden.entrySet()) {
												deviceJson.add(entry8.getKey(), entry8.getValue());
											}
											deviceJson.add(entry7.getKey(), entry7.getValue());
										}
										deviceJson.add(entry6.getKey(), entry6.getValue());
									}
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

			String CR27ShipURL = InitialNode + deviceJson;
			System.out.println(CR27ShipURL);

			mzexecutor = (JavascriptExecutor) ChDriver;

			ChDriver.get(CR27ShipURL);
			Thread.sleep(5000);
			getscreenshot();

			String AB = ChDriver.findElement(By.xpath("/html/body/pre")).getText();
			// System.out.println("AB : " + AB);

			String[] Result1 = AB.split("Label");
			String FinalResult = Result1[0];

			System.out.println("FinalResult : " + FinalResult);

			File src1 = new File("./DataFile/CR27SHIPJSON.xls");
			FileOutputStream fis1 = new FileOutputStream(src1);
			Sheet sh2 = workbook.getSheet("Sheet1");
			sh2.getRow(f).createCell(61).setCellValue(FinalResult);
			workbook.write(fis1);

			// close the file
			fis1.close();

			// Check Expected and Act Result
			String ExpResult = formatter.formatCellValue(sh1.getRow(f).getCell(60));

			// SELECT * FROM WebServiceErrorCodeMessage

			if (!FinalResult.contains(ExpResult)) {
				File src1srcstatus = new File("./DataFile/CR27SHIPJSON.xls");
				FileOutputStream fis2 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(62).setCellValue("FAIL");
				workbook.write(fis2);

				// close the file
				fis2.close();
			} else {
				File src1srcstatus = new File("./DataFile/CR27SHIPJSON.xls");
				FileOutputStream fis3 = new FileOutputStream(src1srcstatus);
				Sheet sh2srcstatus = workbook.getSheet("Sheet1");
				sh2srcstatus.getRow(f).createCell(62).setCellValue("PASS");
				workbook.write(fis3);

				// close the file
				fis3.close();
			}
		}

		msg.append("\n" + "Please Find attached sheet for CR27 Ship URL Result" + "\n\n");
		msg.append("*** This is automated generated email and send through automation script" + "\n");

		String subject = "Selenium Automation Script: CR27 Ship URL Process";
		try {
			Email.sendMail("ravina.prajapati@samyak.com,asharma@samyak.com,parth.doshi@samyak.com", subject, msg.toString(), "./DataFile/CR27SHIPJSON.xls");
		} catch (Exception ex) {
			Logger.getLogger(CR27ShipURL.class.getName()).log(Level.SEVERE, null, ex);
		}

		System.out.println("The End !!!!!!!!!!!!");
	}

	public static void getscreenshot() throws Exception {
		File scrFile = ((TakesScreenshot) ChDriver).getScreenshotAs(OutputType.FILE);
		// The below method will save the screen shot in d drive with name
		// "screenshot.png"

		int s = f;

		FileUtils.copyFile(scrFile, new File(".\\Ship_Screenshot\\" + "Result" + s + ".png"));

	}
}