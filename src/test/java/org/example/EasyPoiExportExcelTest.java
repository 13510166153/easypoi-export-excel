package org.example;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import org.example.dto.CaseDescDTO;
import org.example.dto.CaseDetailDTO;
import org.example.utils.ExcelBeanUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

/**
 * 描述:easypoi实现
 *
 * @author Zhanggq
 * @date 2020/6/3
 */
public class EasyPoiExportExcelTest {

    private static final Logger log = LoggerFactory.getLogger(EasyPoiExportExcelTest.class);
    private static final String dateFormat = "yyyy-MM-dd";

    private static final SimpleDateFormat simpleDateFormat = new SimpleDateFormat(dateFormat);

    @Test
    public void testExport() {
        try {
            List<CaseDescDTO> caseDescDTOS = listCaseDescDTO();
            List<CaseDetailDTO> caseDetailDTOS = listCaseDetailDTO();
            ExportParams exportParams = new ExportParams();
            exportParams.setSheetName("我是sheet名字");
            // 生成workbook 并导出
            Workbook workbook = ExcelExportUtil.exportExcel(exportParams, CaseDescDTO.class, caseDescDTOS);
            Sheet firstSheet = workbook.getSheetAt(0);
            int lastRowNum = firstSheet.getLastRowNum();

            Row newFirtRow = firstSheet.createRow(lastRowNum++);
            String[] newFirstRowHeader = new String[]{"序号", "案件名称", "案件状态", "在办专家名单（绿色表示专家已接受，黄色表示等待专家接受，红色表示专家超时未接受）", "案件发布时间", "承办人和电话号码"};
            for (int i = 0; i < newFirstRowHeader.length; i++) {
                Cell tempCell = newFirtRow.createCell(i);
                tempCell.setCellValue(newFirstRowHeader[i]);
            }
            List<Map<Integer, Object>> dataList = ExcelBeanUtil.manageCaseDetailList(caseDetailDTOS);
            int size = dataList.size();
            int rowIndex = lastRowNum;
            Row row;
            Object obj;
            //合并的单元格
            List<CellRangeAddress> craList=new ArrayList<>();
            for (Map<Integer, Object> rowMap : dataList) {
                try {
                    row = firstSheet.createRow(rowIndex++);
                    //获取当前行每一个单元格不相同的内容
                    for (int i = 0; i < newFirstRowHeader.length; i++) {
                        obj = rowMap.get(i);
                        if (obj == null) {
                            row.createCell(i).setCellValue("");
                        } else if (obj instanceof Date) {
                            String tempDate = simpleDateFormat.format((Date) obj);
                            row.createCell(i).setCellValue((tempDate == null) ? "" : tempDate);
                        } else {
                            System.out.println(obj);
                            row.createCell(i).setCellValue(String.valueOf(obj));
                        }
                    }
                } catch (Exception e) {
                    log.debug("excel sheet填充数据 发生异常： ", e.fillInStackTrace());
                }
            }


            // TODO: 2020/6/3 合并内容相同的行
            //输出Excel文件
            File savefile = new File("D:\\民行\\统计功能\\");
            if (!savefile.exists()) {
                boolean result = savefile.mkdirs();
                System.out.println("目录不存在，创建" + result);
            }
            FileOutputStream fos = new FileOutputStream("D:\\民行\\统计功能\\EasyPoi导出Excel.xls");
            workbook.write(fos);
            fos.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public List<CaseDescDTO> listCaseDescDTO() {
        List<CaseDescDTO> caseDescDTOS = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            CaseDescDTO caseDescDTO = new CaseDescDTO();
            caseDescDTO.setCaseTotalDesc("案件咨询总体情况描述" + i);
            caseDescDTO.setStatisticalDate("2020-05-2" + i);
            caseDescDTOS.add(caseDescDTO);
        }
        return caseDescDTOS;
    }

    public List<CaseDetailDTO> listCaseDetailDTO() {
        List<CaseDetailDTO> caseDetailDTOS = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            for (int j = 0; j < 3; j++) {
                CaseDetailDTO caseDetailDTO = new CaseDetailDTO();
                caseDetailDTO.setId(i + "");
                caseDetailDTO.setCaseName("caseName" + i);
                caseDetailDTO.setCaseStatus("caseStatus" + i);
                caseDetailDTO.setCasePublishDate("2020-05-2" + i);
                caseDetailDTO.setProExp("proExp" + j);
                caseDetailDTO.setPromoterAndPhone("承办人和电话号码" + j);
                caseDetailDTOS.add(caseDetailDTO);
            }
        }
        return caseDetailDTOS;
    }
}
