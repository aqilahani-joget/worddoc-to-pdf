package org.joget.sample;

import org.docx4j.Docx4J;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.joget.apps.app.dao.FormDefinitionDao;
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
        String filenameColumn = (String) properties.get("filenameColumn");
        
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
        String docxFilename = row.getProperty(filenameColumn);
        
        LogUtil.info(getClassName(), "filename: " + docxFilename);
        
        // check if file is word doc or not
        if (!docxFilename.endsWith(".docx") && !docxFilename.endsWith(".doc")) {
            LogUtil.info(getClassName(), "File is not a word document (.docx, .doc)");
            return null;
        }
        String pdfFilename = docxFilename.substring(0, docxFilename.lastIndexOf(".")) + ".pdf";
        try {
            File file = new File((path + docxFilename));
            if (file.exists()) {
                WordprocessingMLPackage wordMLPackage = Docx4J.load(file);
                OutputStream os = new FileOutputStream(new File(path + pdfFilename));
                Docx4J.toPDF(wordMLPackage, os);
            }
            LogUtil.info(getClassName(), "File does not exist");
            
        } catch (Docx4JException | FileNotFoundException ex) {
            LogUtil.error(getClassName(), ex, "");
        }
        return null;
    }
}
