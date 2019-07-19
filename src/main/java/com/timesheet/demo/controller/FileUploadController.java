package com.timesheet.demo.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.stream.Collectors;

import com.timesheet.demo.storage.StorageFileNotFoundException;
import com.timesheet.demo.storage.StorageService;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.MvcUriComponentsBuilder;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;


@Controller
public class FileUploadController {

    private final StorageService storageService;

    @Autowired
    public FileUploadController(StorageService storageService) {
        this.storageService = storageService;
    }

    @GetMapping("/")
    public String listUploadedFiles(Model model) throws IOException {

        model.addAttribute("files", storageService.loadAll().map(
                path -> MvcUriComponentsBuilder.fromMethodName(FileUploadController.class,
                        "serveFile", path.getFileName().toString()).build().toString())
                .collect(Collectors.toList()));

        return "uploadForm";
    }

    @GetMapping("/files/{filename:.+}")
    @ResponseBody
    public ResponseEntity<Resource> serveFile(@PathVariable String filename) {

        Resource file = storageService.loadAsResource(filename);
        return ResponseEntity.ok().header(HttpHeaders.CONTENT_DISPOSITION,
                "attachment; filename=\"" + file.getFilename() + "\"").body(file);
    }

    /**
     * @Author zc
     * @ClassName FileUploadController
     * @Date 11:45 AM 7/12/2019
     * @Version 1.0
     * @Description 注意在判断Excel表中单元格是否为空是是判断cell而不是判断cell.getStringCellValue()
     **/

    @PostMapping("/")
    public String handleFileUpload(@RequestParam("file") MultipartFile file, RedirectAttributes redirectAttributes) {

        //isWeekend("2019-12-22");

        // File fileST = new File("C:\\timeSheet\\time Sheeter\\one\\" + file.getOriginalFilename());

        /**创建Excel，读取文件内容；
         *XSSFWorkbook workbookST = new XSSFWorkbook(FileUtils.openInputStream(fileST));
         *读取想要处理的工作表sheet；
         *Sheet sheet = workbookST.getSheetAt(1);
         * */

        File fileNew = new File("C:\\timeSheet\\time Sheeter\\one\\title.xlsx");

        /**新建工作区；
         *创建Excel工作簿；
         *XSSFWorkbook workbook = new XSSFWorkbook();
         *创建一个工作表sheet；
         *Sheet sheetNew = workbook.createSheet();
         *创建第一行;
         *Row row = sheet.createRow(n);
         **/

        try {
            //写文件
            /**创建Excel工作簿*/
            XSSFWorkbook workbook = new XSSFWorkbook();
            /**创建单元格样式*/
            //XSSFCellStyle style = workbook.createCellStyle();
            //style.setFillBackgroundColor(XSSFColor.toXSSFColor(new XSSFColor()));
            /**创建一个工作表sheet*/
            Sheet sheetNew = workbook.createSheet();
            /**创建一个新的Cell对象*/
            Cell cellNew = null;
            //读文件
            /**创建XSSFWorkbook对象，读取文件内容*/
            XSSFWorkbook workbookST = new XSSFWorkbook(file.getInputStream());
            /**读取想要处理的工作表sheet*/
            Sheet sheet = workbookST.getSheetAt(1);
            /**获取读取文件的行数最大值，也就是有多少行*/
            int lastRowNum = sheet.getLastRowNum();
            /**创建Excel的行对象*/
            Row row = null;
            /**定义一个整型变量变量，用于跳过读取文件空白的第一行和第二行*/
            int w = 0;
            String title = null;
            int remark = 0, roll = 0;
            /**对读取文件中的每一个非空行进行读取*/
            for (int j = 0; j < lastRowNum; j++) {
                /**将行号为1和行号为2的空白行筛选出来*/
                if (j == 0) {
                    w = j;
                } else if (w < lastRowNum) {
                    w = j + 2;
                }
                /**初始化行对象*/
                row = sheet.getRow(w);
                /**获取读取的Excel中每一行中最多的单元格数*/
                int lastCellNum = row.getLastCellNum();
                /**为新建的工作区来创建行对象*/
                Row nextrow = sheetNew.createRow(j);
                /**对每一个单元格进行遍历*/
                if (j == 0) {
                    for (int v = 0; v < lastCellNum; v++) {
                        row.getCell(v).setCellType(Cell.CELL_TYPE_STRING);
                        if ("Roll-off Date".equals(row.getCell(v).getStringCellValue())) {
                            roll = v;
                        } else if ("remarks".equals(row.getCell(v).getStringCellValue())) {
                            remark = v;
                        }
                    }
                }
                for (int i = 0; i < lastCellNum; i++) {
                    /**创建单元格对象*/
                    Cell cell = row.getCell(i);
                    if (null != row.getCell(i)) {
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        /*if ((null != row.getCell(6)) && (null != row.getCell(7) && (null != row.getCell(8)))) {
                            row.getCell(6).setCellType(Cell.CELL_TYPE_STRING);
                            row.getCell(7).setCellType(Cell.CELL_TYPE_STRING);
                            row.getCell(8).setCellType(Cell.CELL_TYPE_STRING);
                            if (!row.getCell(6).getStringCellValue().equals(row.getCell(7).getStringCellValue()) && !row.getCell(7).getStringCellValue().equals(row.getCell(8).getStringCellValue())) {*/

                        /**创建新的cell对象，为新文件增加单元格*/
                        cellNew = nextrow.createCell(i);
                        if (j == 0 && i > remark && i < lastCellNum) {
                            title = cell.getStringCellValue();
                            if ("".equals(title)) {
                                title = "0";
                            }
                            int days = Integer.parseInt(title);
                            Calendar calendar = new GregorianCalendar(1900, 0, -1);
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/YYYY");
                            Date d = calendar.getTime();
                            title = sd.format(DateUtils.addDays(d, Integer.valueOf(days)));
                            cellNew.setCellValue(title);
                        } else if (j > 0 && i == roll + 1) {
                            if ("".equals(title)) {
                                title = "0";
                            } else if ("C".equals(title)) {
                                System.out.println("第" + j + "行" + "第" + i + "列的Cell为值为C");
                                title = "0";
                            }
                            int days = Integer.parseInt(title);
                            Calendar calendar = new GregorianCalendar(1900, 0, -1);
                            SimpleDateFormat sd = new SimpleDateFormat("MM/dd/YYYY");
                            Date d = calendar.getTime();
                            title = sd.format(DateUtils.addDays(d, Integer.valueOf(days)));
                            cellNew.setCellValue(title);
                        } else {
                            /**获取每个单元格的转化后的字符串值*/
                            title = cell.getStringCellValue();
                            /**给新增加的单元格传值*/
                            cellNew.setCellValue(title);
                            /**设置不同内容字的颜色*/
                        }
                    } else {
                        cellNew.setCellValue("凉凉");
                        //System.out.println("第" + j + "行" + "第" + i + "列的Cell为空！");
                    }
                }
            }
            /**创建新的文件*/
            fileNew.createNewFile();
            /**将Excel内容存盘*/
            FileOutputStream stream = FileUtils.openOutputStream(fileNew);
            /**将新建的工作区写入文件中*/
            workbook.write(stream);
            stream.close();
            /**校验ppm与ST,同时是以ST的数据为准，即基于ST的数据比较pmm有什么不同*/
            checkPpmAndST();
            /**校验MyTe与ST,同时是以ST的数据为准，即基于ST的数据比较MyTe有什么不同*/
            checkTeAndST();
        } catch (IOException e) {
            e.printStackTrace();
        }
        storageService.store(file);
        redirectAttributes.addFlashAttribute("message",
                "You successfully uploaded " + file.getOriginalFilename() + "!");
        return "redirect:/";
    }

    private String checkPpmAndST() throws IOException {
        File filePPM = new File("C:\\timeSheet\\time Sheeter\\one\\Apr of PPM.xlsx");

        File filePpmNew = new File("C:\\timeSheet\\time Sheeter\\one\\titlePPM.xlsx");
        try {
            //写文件
            /**创建Excel工作簿*/
            XSSFWorkbook workbook = new XSSFWorkbook();
            /**创建单元格样式*/
            //XSSFCellStyle style = workbook.createCellStyle();
            //style.setFillBackgroundColor(XSSFColor.toXSSFColor(new XSSFColor()));
            /**创建一个工作表sheet*/
            Sheet sheetPpmNew = workbook.createSheet();
            /**创建一个新的Cell对象*/
            Cell cellPpmNew = null;
            XSSFWorkbook workbookPPM = new XSSFWorkbook(FileUtils.openInputStream(filePPM));
            /**读取想要处理的工作表sheet*/
            Sheet sheetPPM = workbookPPM.getSheetAt(1);
            /**获取读取文件的行数最大值，也就是有多少行*/
            int lastRowNumPPM = sheetPPM.getLastRowNum();
            /**创建Excel的行对象*/
            for (int m = 0; m < lastRowNumPPM; m++) {
                Row row = sheetPPM.getRow(m);
                Row nextPpmRow = sheetPpmNew.createRow(m);
                int lastCellNumPPM = row.getLastCellNum();
                for (int n = 0; n < lastCellNumPPM; n++) {
                    Cell cell = row.getCell(n);
                    if (null != cell && null != row.getCell(n)) {
                        /**设置单元格的类型*/
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        /**获取每个单元格的转化后的字符串值*/
                        String titlePPM = cell.getStringCellValue();
                        /**创建新的cell对象，为新文件增加单元格*/
                        cellPpmNew = nextPpmRow.createCell(n);
                        /**给新增加的单元格传值*/
                        cellPpmNew.setCellValue(titlePPM);
                    }
                }
            }
            /**创建新的文件*/
            filePpmNew.createNewFile();
            /**将Excel内容存盘*/
            FileOutputStream stream = FileUtils.openOutputStream(filePpmNew);
            /**将新建的工作区写入文件中*/
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    private String checkTeAndST() {
        File fileTE = new File("C:\\timeSheet\\time Sheeter\\one\\Apr of MY Time Details.xlsx");

        File fileTeNew = new File("C:\\timeSheet\\time Sheeter\\one\\titleTE.xlsx");
        try {
            //写文件
            /**创建Excel工作簿*/
            XSSFWorkbook workbook = new XSSFWorkbook();
            /**创建单元格样式*/
            //XSSFCellStyle style = workbook.createCellStyle();
            //style.setFillBackgroundColor(XSSFColor.toXSSFColor(new XSSFColor()));
            /**创建一个工作表sheet*/
            Sheet sheetTeNew = workbook.createSheet();
            /**创建一个新的Cell对象*/
            Cell cellTeNew = null;
            XSSFWorkbook workbookTE = new XSSFWorkbook(FileUtils.openInputStream(fileTE));
            /**读取想要处理的工作表sheet*/
            Sheet sheetTE = workbookTE.getSheetAt(1);
            /**获取读取文件的行数最大值，也就是有多少行*/
            int lastRowNumTE = sheetTE.getLastRowNum();
            /**创建Excel的行对象*/
            for (int z = 0; z < lastRowNumTE; z++) {
                Row row = sheetTE.getRow(z);
                Row nextTeRow = sheetTeNew.createRow(z);
                int lastCellNumTE = row.getLastCellNum();
                for (int c = 0; c < lastCellNumTE; c++) {
                    Cell cell = row.getCell(c);
                    if (null != cell && null != row.getCell(c)) {
                        /**设置单元格的类型*/
                        cell.setCellType(Cell.CELL_TYPE_STRING);
                        /**获取每个单元格的转化后的字符串值*/
                        String titlePPM = cell.getStringCellValue();
                        /**创建新的cell对象，为新文件增加单元格*/
                        cellTeNew = nextTeRow.createCell(c);
                        /**给新增加的单元格传值*/
                        cellTeNew.setCellValue(titlePPM);
                    }
                }
            }
            /**创建新的文件*/
            fileTeNew.createNewFile();
            /**将Excel内容存盘*/
            FileOutputStream stream = FileUtils.openOutputStream(fileTeNew);
            /**将新建的工作区写入文件中*/
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    @GetMapping("/date")
    public static String isWeekend(@RequestParam("bDate") String bDate) throws ParseException {
        DateFormat format1 = new SimpleDateFormat("MM/dd/YYYY");
        Date bdate = format1.parse(bDate);
        Calendar cal = Calendar.getInstance();
        cal.setTime(bdate);
        if (cal.get(Calendar.DAY_OF_WEEK) == Calendar.SATURDAY || cal.get(Calendar.DAY_OF_WEEK) == Calendar.SUNDAY) {
            // return "OK";
            System.out.println("今天属于周末，不用上班！");
        } else {
            // return "NO";
            System.out.println("今天是工作日，要开始混底薪了！");
        }

        return null;
    }


    @ExceptionHandler(StorageFileNotFoundException.class)
    public ResponseEntity<?> handleStorageFileNotFound(StorageFileNotFoundException exc) {
        return ResponseEntity.notFound().build();
    }
}
