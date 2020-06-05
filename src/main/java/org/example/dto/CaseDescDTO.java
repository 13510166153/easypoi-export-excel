package org.example.dto;

import cn.afterturn.easypoi.excel.annotation.Excel;
import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;

import java.io.Serializable;

/**
 * 描述:
 *
 * @author Zhanggq
 * @date 2020/6/3
 */
@Data
@ApiModel(value = "案件咨询总体情况描述=案件咨询总体情况描述第一行+案件咨询总体情况描述其他行")
public class CaseDescDTO implements Serializable {
    private static final long serialVersionUID = 1L;
        @ApiModelProperty(value = "统计日期")
        @Excel(name = "统计日期", height = 20, width = 30)
        private String statisticalDate;

        @ApiModelProperty(value = "案件咨询总体情况描述")
        @Excel(name = "案件咨询总体情况描述", height = 20, width = 100)
        private String caseTotalDesc;
}
