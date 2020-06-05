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
@ApiModel(value = "案件详情")
public class CaseDetailDTO implements Serializable {
    private static final long serialVersionUID = 1L;
    @ApiModelProperty(value = "序号")
    @Excel(name = "序号",mergeVertical = true)
    private String id;

    @ApiModelProperty(value = "案件名称")
    @Excel(name = "案件名称",mergeVertical = true)
    private String caseName;

    @ApiModelProperty(value = "案件状态")
    @Excel(name = "案件状态",mergeVertical = true)
    private String caseStatus;

    @ApiModelProperty(value = "案件发布时间")
    @Excel(name = "案件发布时间",mergeVertical = true)
    private String casePublishDate;

    @ApiModelProperty(value = "在办专家名单（绿色表示专家已接受，黄色表示等待专家接受，红色表示专家超时未接受）")
    @Excel(name = "在办专家名单（绿色表示专家已接受，黄色表示等待专家接受，红色表示专家超时未接受）")
    private String proExp;

    @ApiModelProperty(value = "承办人和电话号码")
    @Excel(name = "承办人和电话号码")
    private String promoterAndPhone;

}
