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
File file1 = new File("file1.xlsx");
File file2 = new File("file2.xlsx");
File file3 = new File("file3.xlsx");

//กำหนดว่าจะให้ sheet ต่างๆ ใน excel file เดิม ตอนที่รวมใน file ใหม่แล้วมีชื่อ sheet ว่าอะไรบ้าง
//file นี้มี 2 sheets เลยเปลี่ยนชื่อทั้ง 2 sheets
ExcelFile exFile1 = ExcelFile.from(file1).andWithSheetName("ชื่อ sheet 1").andWithSheetName("ชื่อ sheet 2");
//file นี้มี sheet เดียว
ExcelFile exFile2 = ExcelFile.from(file2).andWithSheetName("ชื่อ sheet 3");
//file นี้มี sheet เดีย
ExcelFile exFile3 = ExcelFile.from(file3).andWithSheetName("ชื่อ sheet 4");

//excel file ปลายทาง ตอนที่รวมเสร็จแล้ว
//ไม่ต้องมีอยู่ก่อนแล้วก็ได้
File output1 = new File("output.xlsx");

//การเรียกใช้
File integratedFile = ExcelSheetIntegrator.newInstance()
		.addExcelFile(exFile1)
		.addExcelFile(exFile2)
		.addExcelFile(exFile3)
		.toTargetFile(output1)
		.integrate();
```
##จากนั้น
Sheet จาก file1.xlsx, file2.xlsx และ file3.xlsx จะถูกเอามารวมอยู่ใน file output.xlsx file เดียว

##เอามาใช้แก้ปัญหาอะไร
พอดีผมทำ jasper report แล้วมันไม่สามารถทำ multiple sheets ได้  คือแต่ละ sheet มีข้อมูลที่ต่างกันน่ะ ข้อมูลคนละแบบกันเลย

ลูกค้าอยากได้ ข้อมูลต่างๆ รวมอยู่ใน file เดียวกัน
ผมก็เลยใช้ iReport ทำ report แต่ละแบบ  แต่ละ sheet แต่ละ file ตามปกติ
จากนั้นก็ใช้ lib ที่ผมเขียนขึ้นมานี้ เพื่อ integrate sheet ต่างๆ ให้รวมอยู่ใน file เดียวครับ
แล้วค่อย export ออกมาให้ลูกค้าครับ

หวังว่าจะเป็นประโชน์ ไม่มากก็น้อย  สำหรับคนที่จะเอาไปใช้ต่อน่ะครับ
ถ้ามีอะไรผิดพลาด  ช่วยแจ้ง bug ให้ด้วยน่ะครับ  แล้วมีจะรีบแก้ไขโดยเร็วที่สุด  ขอบคุณครับ
