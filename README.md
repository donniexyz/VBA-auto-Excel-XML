# VBA-auto-Excel-XML

## Converting flat excel table/data into structured XML text

Automatically converting this normal excel table

action.name|action.class|result.name|result.nodeTypedValue
-----|------|------|-------
exposure|com.doztronics.web.controller.ExposureAction|detail|exposure_detail.jsp
exposure|com.doztronics.web.controller.ExposureAction|edit|exposure_edit.jsp
exposure|com.doztronics.web.controller.ExposureAction|manual_add|exposure_manual_add.jsp
exposure|com.doztronics.web.controller.ExposureAction|manual_edit|exposure_manual_edit.jsp
exposure|com.doztronics.web.controller.ExposureAction|manual_detail|exposure_manual_detail.jsp
exposure|com.doztronics.web.controller.ExposureAction|bic_edit|exposure_bic_edit.jsp
run|com.doztronics.web.controller.RunAction|cbi_edit|run_cbi_edit.jsp
run|com.doztronics.web.controller.RunAction|cci_edit|run_cci_edit.jsp
run|com.doztronics.web.controller.RunAction|search|run_search.jsp
run|com.doztronics.web.controller.RunAction|verifying_list|run_verifying_list.jsp
run|com.doztronics.web.controller.RunAction|pop_up_close|run_pop_up_close.jsp
run|com.doztronics.web.controller.RunAction|customer_summary|run_customer_summary.jsp



into this XML

```
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE struts PUBLIC
        "-//Apache Software Foundation//DTD Struts Configuration 2.3//EN"
        "http://struts.apache.org/dtds/struts-2.3.dtd">
<struts>

    <package>
        <action name="exposure" class="com.doztronics.web.controller.ExposureAction">
            <result name="detail">exposure_detail.jsp</result>
            <result name="edit">exposure_edit.jsp</result>
            <result name="manual_add">exposure_manual_add.jsp</result>
            <result name="manual_edit">exposure_manual_edit.jsp</result>
            <result name="manual_detail">exposure_manual_detail.jsp</result>
            <result name="bic_edit">exposure_bic_edit.jsp</result>
        </action>
        <action name="run" class="com.doztronics.web.controller.RunAction">
            <result name="cbi_edit">run_cbi_edit.jsp</result>
            <result name="cci_edit">run_cci_edit.jsp</result>
            <result name="search">run_search.jsp</result>
            <result name="verifying_list">run_verifying_list.jsp</result>
            <result name="pop_up_close">run_pop_up_close.jsp</result>
            <result name="customer_summary">run_customer_summary.jsp</result>
        </action>
    </package>
</struts>
```
