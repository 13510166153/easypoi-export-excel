package org.example.utils;


import org.example.dto.CaseDetailDTO;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 导入excel bean数据工具类
 */
public class ExcelBeanUtil {


    /**
     * 处理案件详情
     *
     * @param caseDetailDTOS
     * @return
     */
    public static List<Map<Integer, Object>> manageCaseDetailList(final List<CaseDetailDTO> caseDetailDTOS) {
        List<Map<Integer, Object>> dataList = new ArrayList<>();
        if (caseDetailDTOS != null && caseDetailDTOS.size() > 0) {
            int length = caseDetailDTOS.size();
            Map<Integer, Object> dataMap;
            for (int i = 0; i < length; i++) {
                CaseDetailDTO caseDetailDTO = caseDetailDTOS.get(i);
                dataMap = new HashMap<>();
                dataMap.put(0, caseDetailDTO.getId());
                dataMap.put(1, caseDetailDTO.getCaseName());
                dataMap.put(2, caseDetailDTO.getCaseStatus());
                dataMap.put(3,caseDetailDTO.getProExp());
                dataMap.put(4,caseDetailDTO.getCasePublishDate());
                dataMap.put(5,caseDetailDTO.getPromoterAndPhone());
                dataList.add(dataMap);
            }
        }
        return dataList;
    }
}