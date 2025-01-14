package org.joget.sample;

import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;

import java.util.HashMap;
import java.util.Map;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import java.io.IOException;
import org.joget.apps.app.dao.FormDefinitionDao;
import org.joget.apps.app.lib.DatabaseUpdateTool;
import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.model.FormDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.service.FileUtil;
import org.joget.commons.util.LogUtil;
import org.joget.plugin.base.DefaultApplicationPlugin;
import org.joget.workflow.model.WorkflowAssignment;

/**
 *
 * @author Default
 */
public class WordDocToPdfProcessTool extends DefaultApplicationPlugin {
    private final static String MESSAGE_PATH = "messages/wordDocToPdfProcessTool";
    
    @Override
    public String getName() {
        return AppPluginUtil.getMessage("org.joget.sample.WordDocToPdf.title", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getVersion() {
        return "1.0.0";
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("org.joget.sample.WordDocToPdf.desc", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("org.joget.sample.WordDocToPdf.title", getClassName(), MESSAGE_PATH);
    }

    @Override
    public String getClassName() {
        return this.getClass().getName();
    }

    @Override
    public String getPropertyOptions() {
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        String appId = appDef.getId();
        String appVersion = appDef.getVersion().toString();
        Object[] arguments = new Object[]{getLabel(), appId, appVersion};
        String json = AppUtil.readPluginResource(getClass().getName(), "/properties/wordDocToPdfProcessTool.json", arguments, true, MESSAGE_PATH);
        return json;
    }
    
    @Override
    public Object execute(Map properties) {
        
        //Get FormDefId from properties
        String formDefId = (String) properties.get("formDefId");

        //Get record Id from process
        WorkflowAssignment wfAssignment = (WorkflowAssignment) properties.get("workflowAssignment");
        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");
        String id = appService.getOriginProcessId(wfAssignment.getProcessId());

        LogUtil.info(getClassName(), "recordId: " + id);
        String docFilenameColumn = (String) properties.get("docFilenameColumn");
        String pdfFilenameColumn = (String) properties.get("pdfFilenameColumn");
        
        //Load the original Form Data record
        AppDefinition appDef = (AppDefinition) properties.get("appDef");
        FormRow row = new FormRow();
        FormRowSet rowSet = appService.loadFormData(appDef.getAppId(), appDef.getVersion().toString(), formDefId, id);
        if (!rowSet.isEmpty()) {
            row = rowSet.get(0);
        }
        
        // get form upload path
        FormDefinitionDao formDefinitionDao = (FormDefinitionDao) AppUtil.getApplicationContext().getBean("formDefinitionDao");
        FormDefinition formDef = formDefinitionDao.loadById(formDefId, appDef);
        String path = FileUtil.getUploadPath(formDef.getTableName(), id);
        
        LogUtil.info(getClassName(), "path: " + path);
        String docxFilename = row.getProperty(docFilenameColumn);
        
        LogUtil.info(getClassName(), "filename: " + docxFilename);
        
        // check if file is word doc or not
        if (!docxFilename.endsWith(".docx") && !docxFilename.endsWith(".doc")) {
            LogUtil.info(getClassName(), "File is not a word document (.docx, .doc)");
            return null;
        }
        String pdfFilename = docxFilename.substring(0, docxFilename.lastIndexOf(".")) + ".pdf";
        try (InputStream docFile = new FileInputStream(new File(path + docxFilename));
             OutputStream out = new FileOutputStream(new File(path + pdfFilename))) {
            
            XWPFDocument doc = new XWPFDocument(docFile);
            PdfOptions pdfOptions = PdfOptions.create();
            PdfConverter.getInstance().convert(doc, out, pdfOptions);
            
        } catch (XWPFConverterException | IOException ex) {
            LogUtil.error(getClassName(), ex, "Error converting DOCX to PDF: " + ex.getMessage());
        }

        //update pdf file column
        updatePdfField(wfAssignment, id, pdfFilenameColumn, pdfFilename, formDef.getTableName());

        return null;
    }

    private void updatePdfField(WorkflowAssignment wfAssignment, String recordId, String pdfColumn, String pdfFilename, String tableName) {        
        Map propertiesMap = new HashMap();
        propertiesMap.put("workflowAssignment", wfAssignment);
        propertiesMap.put("jdbcDatasource", "default");
        propertiesMap.put("query", "UPDATE app_fd_" + tableName + " SET c_" + pdfColumn + " = '" + pdfFilename + "' WHERE id = '" + recordId + "'");
        new DatabaseUpdateTool().execute(propertiesMap);
    }
}
