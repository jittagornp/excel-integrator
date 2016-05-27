# excel-integrator

ไวสำหรับรวม excel ไฟล์ (sheets) อื่นๆ ให้เป็น excel ไฟล์เดียว
```
File mergedFile = ExcelSheetIntegrator.newInstance()
                .addExcelFile(exFile1)
                .addExcelFile(exFile2)
                .addExcelFile(exFile3)
                .toTargetFile(output1)
                .merge();
```
##example
```
File input1 = new File("file1.xlsx");
File input2 = new File("file2.xlsx");
File input3 = new File("file3.xlsx");

ExcelFile exFile1 = ExcelFile.fromFile(input1).withSheetName("ชื่อ sheet 1");
ExcelFile exFile2 = ExcelFile.fromFile(input2).withSheetName("ชื่อ sheet 2");
ExcelFile exFile3 = ExcelFile.fromFile(input3).withSheetName("ชื่อ sheet 3");

File output1 = new File("output.xlsx");

File mergedFile = ExcelSheetIntegrator.newInstance()
		.addExcelFile(exFile1)
		.addExcelFile(exFile2)
		.addExcelFile(exFile3)
		.toTargetFile(output1)
		.merge();
```
