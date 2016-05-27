# excel-integrator

ไว้สำหรับรวม excel ไฟล์ (sheets) อื่นๆ ให้เป็น excel ไฟล์เดียว
```java
File integratedFile = ExcelSheetIntegrator.newInstance()
                .addExcelFile(exFile1)
                .addExcelFile(exFile2)
                .addExcelFile(exFile3)
                .toTargetFile(output1)
                .integrate();
```
##example
```java
//excel file ต่างๆ ที่อยากเอามารวม
File input1 = new File("file1.xlsx");
File input2 = new File("file2.xlsx");
File input3 = new File("file3.xlsx");

//กำหนดว่าจะให้ sheet ต่างๆ ใน excel file เดิม ตอนที่รวมใน file ใหม่แล้วมีชื่อ sheet ว่าอะไรบ้าง
ExcelFile exFile1 = ExcelFile.fromFile(input1).withSheetName("ชื่อ sheet 1").withSheetName("ชื่อ sheet 2");
ExcelFile exFile2 = ExcelFile.fromFile(input2).withSheetName("ชื่อ sheet 3");
ExcelFile exFile3 = ExcelFile.fromFile(input3).withSheetName("ชื่อ sheet 4");

//excel file ปลายทาง ตอนที่รวมเสร็จแล้ว
File output1 = new File("output.xlsx");

//การเรียกใช้
File integratedFile = ExcelSheetIntegrator.newInstance()
		.addExcelFile(exFile1)
		.addExcelFile(exFile2)
		.addExcelFile(exFile3)
		.toTargetFile(output1)
		.integrate();
```
