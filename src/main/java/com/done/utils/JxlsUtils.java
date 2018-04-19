package com.done.utils;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.done.exception.CustomException;
import org.apache.commons.lang3.StringUtils;
import org.jxls.common.Context;
import org.jxls.expression.JexlExpressionEvaluator;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;

import javax.servlet.http.HttpServletResponse;

/**
 * 自定义excel模板导出excel
 */
public class JxlsUtils{

    public static void exportExcel(InputStream is, OutputStream os, Map<String, Object> model) throws IOException{
        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }
        JxlsHelper jxlsHelper = JxlsHelper.getInstance();
        Transformer transformer  = jxlsHelper.createTransformer(is, os);
        JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator)transformer.getTransformationConfig().getExpressionEvaluator();
        Map<String, Object> funcs = new HashMap<String, Object>();
        funcs.put("utils", new JxlsUtils());    //添加自定义功能
        evaluator.getJexlEngine().setFunctions(funcs);
        jxlsHelper.processTemplate(context, transformer);
    }

    /**
     * <p>
     * 自定义excel模板导出excel-自定义路径导出<br>
     * </p>
     * <p>
     * -----版本-----变更日期-----责任人-----变更内容<br>
     * ─────────────────────────────────────<br>
     * V1.0.0 2018年04月19日 luoxiang 初版<br>
     *
     * @param templateExcelURL
     *            excel模板路径(包含至excel文件名)
     * @param customeExcelName
     *            自定义导出的excel文件名称(默认名称:Excel+年月日时分秒)
     * @param list
     *            导出数据的List<Map<String, Object>>, 不批量导出时可传入null,excel中取值遍历示例:
     *            jx:each(items="list" var = "temp" lastCell = "H12")->H12是指A12-H12行往下遍历
     *            jx:area(lastCell="H13")->H13是指表格范围A1-H13
     * @param param
     *            导出数据的Map<String, Object>,可传入null,
     *            excel中取值时示例(${param.参数名称}):${param.name}
     * @param response
     *            HttpServletResponse
     * @throws CustomException
     *             自定义异常,成功无返回,无异常 void,失败异常返回错误原因(e.getMessage())
     * @since XMJR V3.0.0
     *        </p>
     */
    public static void exportExcel(String templateExcelURL, String customeExcelName,List<Map<String, Object>> list, Map<String, Object> param,
                                   HttpServletResponse response) throws CustomException {
        if (StringUtils.isBlank(templateExcelURL)) {
            throw new CustomException("Excel模板URL不能为空");
        }

        File file = new File(templateExcelURL);

        if (file.isDirectory()) {
            throw new CustomException("Excel导出模板路径需包含模板文件名称");
        }

        if (!file.exists()) {
            throw new CustomException("Excel导出模板文件不存在,请联系管理人员添加模板");
        }

        if (StringUtils.isBlank(customeExcelName)) {
            try {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
                customeExcelName = "Excel" + simpleDateFormat.format(new Date()) + ".xls";
            } catch (Exception e) {
                throw new CustomException("Excel新建默认名称日期转换异常,请重试");
            }
        } else {
            if (!customeExcelName.contains(".xls")) {
                customeExcelName += ".xls";
            }
        }

        Map<String,Object> model = new HashMap<String,Object>();
        if (list != null && !list.isEmpty()) {
            model.put("list", list);
        }
        if (param != null && !param.isEmpty()) {
            model.put("param", param);
        }
        if (model.isEmpty()) {
            throw new CustomException("Excel导出数据为空,请传入导出数据");
        }

        if (model.isEmpty()) {
            throw new CustomException("Excel导出数据为空,请传入导出数据");
        }

        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }

        InputStream inputStream = null;
        OutputStream outputStream = null;
        System.out.println("***************导入Excel模板的数据***************"+model);
        try {
            // 设置响应
            response.setHeader("Content-Disposition", "attachment;filename="
                    + new String(customeExcelName.getBytes("GBK"), "ISO8859-1"));
            response.setContentType("application/vnd.ms-excel");

            //将模板放在项目中
            //InputStream in = ExportExcel.class.getResourceAsStream("/template/" + templateName);
            inputStream = new BufferedInputStream(new FileInputStream(templateExcelURL));
            outputStream = response.getOutputStream();
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            Transformer transformer  = jxlsHelper.createTransformer(inputStream, outputStream);

            /*-------------------- 给模板添加自定义功能 ----------------------------*/
            JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator)transformer.getTransformationConfig().getExpressionEvaluator();
            Map<String, Object> funcs = new HashMap<String, Object>();
            funcs.put("utils", new JxlsUtils());    //eg：${utils:dateFmt(date,"yyyy-MM-dd")}-格式化日期
            evaluator.getJexlEngine().setFunctions(funcs);

            jxlsHelper.processTemplate(context, transformer);
            outputStream.flush();
        } catch (Exception e) {
            throw new CustomException("Excel导出异常,请重试");
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    throw new CustomException("Excel导出异常,inputStream.close异常");
                }
            }
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    throw new CustomException("Excel导出异常,outputStream.close异常");
                }
            }
        }
    }

    /**
     * <p>
     * 自定义excel模板导出excel-浏览器导出<br>
     * </p>
     * <p>
     * -----版本-----变更日期-----责任人-----变更内容<br>
     * ─────────────────────────────────────<br>
     * V1.0.0 2018年04月19日 luoxiang 初版<br>
     *
     * @param templateExcelURL
     *            excel模板路径(包含至excel文件名)
     * @param list
     *            导出数据的List<Map<String, Object>>, 不批量导出时可传入null,excel中取值遍历示例:
     *            jx:each(items="list" var = "temp" lastCell = "H12")->H12是指A12-H12行往下遍历
     *            jx:area(lastCell="H13")->H13是指表格范围A1-H13
     * @param param
     *            导出数据的Map<String, Object>,可传入null,
     *            excel中取值时示例(${param.参数名称}):${param.name}
     * @param outPath
     *            指定输出路径全称(***.xls)
     * @throws CustomException
     *             自定义异常,成功无返回,无异常 void,失败异常返回错误原因(e.getMessage())
     * @since XMJR V3.0.0
     *        </p>
     */
    public static void exportExcelToPath(String templateExcelURL,List<Map<String, Object>> list, Map<String, Object> param,
                                         String outPath) throws CustomException{
        if (StringUtils.isBlank(templateExcelURL)) {
            throw new CustomException("Excel模板URL不能为空");
        }

        File file = new File(templateExcelURL);
        File file2 = new File(outPath);


        if (file.isDirectory()) {
            throw new CustomException("Excel导出模板路径需包含模板文件名称");
        }

        if (file2.isDirectory()) {
            throw new CustomException("Excel导出路径需包含文件名称");
        }

        if (!file.exists()) {
            throw new CustomException("Excel导出模板文件不存在,请联系管理人员添加模板");
        }

        Map<String,Object> model = new HashMap<String,Object>();
        if (list != null && !list.isEmpty()) {
            model.put("list", list);
        }
        if (param != null && !param.isEmpty()) {
            model.put("param", param);
        }
        if (model.isEmpty()) {
            throw new CustomException("Excel导出数据为空,请传入导出数据");
        }

        if (model.isEmpty()) {
            throw new CustomException("Excel导出数据为空,请传入导出数据");
        }

        Context context = new Context();
        if (model != null) {
            for (String key : model.keySet()) {
                context.putVar(key, model.get(key));
            }
        }

        InputStream inputStream = null;
        OutputStream outputStream = null;
        System.out.println("***************导入Excel模板的数据***************"+model);
        try {

            //将模板放在项目中
            //InputStream in = ExportExcel.class.getResourceAsStream("/template/" + templateName);
            inputStream = new BufferedInputStream(new FileInputStream(templateExcelURL));
            outputStream =new FileOutputStream(file2);
            JxlsHelper jxlsHelper = JxlsHelper.getInstance();
            Transformer transformer  = jxlsHelper.createTransformer(inputStream, outputStream);

            /*-------------------- 给模板添加自定义功能 ----------------------------*/
            JexlExpressionEvaluator evaluator = (JexlExpressionEvaluator)transformer.getTransformationConfig().getExpressionEvaluator();
            Map<String, Object> funcs = new HashMap<String, Object>();
            //eg：${utils:dateFmt(date,"yyyy-MM-dd")}-格式化日期
            funcs.put("utils", new JxlsUtils());
            evaluator.getJexlEngine().setFunctions(funcs);

            jxlsHelper.processTemplate(context, transformer);
            //jxlsHelper.processTemplate(inputStream,outputStream,context);
            outputStream.flush();
        } catch (Exception e) {
            e.printStackTrace();
            throw new CustomException("Excel导出异常,请重试");
        } finally {
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    throw new CustomException("Excel导出异常,inputStream.close异常");
                }
            }
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    throw new CustomException("Excel导出异常,outputStream.close异常");
                }
            }
        }
    }

    public static void exportExcel(File xls, File out, Map<String, Object> model) throws Exception {
        exportExcel(new FileInputStream(xls), new FileOutputStream(out), model);
    }

    public static void exportExcel(String templateName, OutputStream os, Map<String, Object> model) throws Exception {
        File template = new File(templateName);
        if(template!=null){
            exportExcel(new FileInputStream(template), os, model);
        }
    }

    /**
     * 日期格式化
     * @param date
     * @param fmt
     * @return
     */
    public String dateFmt(Date date, String fmt) {
        if (date == null) {
            return "";
        }
        try {
            SimpleDateFormat dateFmt = new SimpleDateFormat(fmt);
            return dateFmt.format(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "";
    }

    /**
     * 手机号脱敏
     * @param phone
     * @return
     */
    public String phoneFmt(String phone){
        return phone.substring(0,3)+"****"+phone.substring(7,11);
    }

}
