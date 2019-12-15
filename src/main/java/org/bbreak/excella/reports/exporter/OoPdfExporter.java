/*-
 * #%L
 * excella-pdfexporter
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.reports.exporter;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.document.DefaultDocumentFormatRegistry;
import org.jodconverter.document.DocumentFamily;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.document.DocumentFormatRegistry;
import org.jodconverter.document.SimpleDocumentFormatRegistry;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * OpenOfficePDF出力エクスポーター
 *
 * @since 1.0
 */
public class OoPdfExporter extends ReportBookExporter {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( OoPdfExporter.class);

    /**
     * 変換タイプ：PDF
     */
    public static final String FORMAT_TYPE = "PDF";

    /**
     * 拡張子：PDF
     */
    public static final String EXTENTION = ".pdf";

    /**
     * OpneOfficeデフォルトポート番号
     */
    public static final int OPENOFFICE_DEFAULT_PORT = 8100;

    /**
     * OpneOfficeポート番号
     */
    private int port = OPENOFFICE_DEFAULT_PORT;

    /**
     * OpneOfficeマネージャ
     */
    private OfficeManager officeManager = null;

    /**
     * OpneOfficeマネージャのコントロール有無
     */
    private boolean controlOfficeManager = false;

    /**
     * コンストラクタ<BR>
     * デフォルトポート番号8100
     */
    public OoPdfExporter() {
    }

    /**
     * コンストラクタ
     *
     * @param port OpneOfficeポート番号
     */
    public OoPdfExporter( int port) {
        this.port = port;
    }

    /**
     * コンストラクタ
     *
     * @param officeManager OpneOfficeマネージャ
     */
    public OoPdfExporter( OfficeManager officeManager) {
        this.officeManager = officeManager;
        controlOfficeManager = true;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.poireports.exporter.ReportBookExporter#output(org.apache.poi.ss.usermodel.Workbook, org.excelparser.BookData, org.poireports.model.ConvertConfiguration)
     */
    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        // TODO POIにより出力されたxlsxファイルはOpenOfficeで読めない不具合あり
        // http://www.openoffice.org/issues/show_bug.cgi?id=97460
        // https://issues.apache.org/bugzilla/show_bug.cgi?id=46419
        // POI3.5より対応済
//        if ( book instanceof XSSFWorkbook) {
//            throw new IllegalArgumentException( "XSSFFile not supported.");
//        }

        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + getFilePath() + "に出力します");
        }

        if ( !controlOfficeManager) {
            officeManager = new DefaultOfficeManagerBuilder().setPortNumber( port).build();
            try {
                officeManager.start();
            } catch ( OfficeException e) {
                throw new ExportException( e);
           }
        }

        File tmpFile = null;
        try {

            OfficeDocumentConverter converter = null;
            if ( configuration.getOptionsProperties().isEmpty()) {
                converter = new OfficeDocumentConverter( officeManager);
            } else {
                DocumentFormatRegistry registry = createDocumentFormatRegistry( configuration);
                converter = new OfficeDocumentConverter( officeManager, registry);
            }

            // 一時フォルダに吐き出し
            ExcelExporter excelExporter = new ExcelExporter();
            tmpFile = File.createTempFile( getClass().getSimpleName(), null);
            String tmpFileName = tmpFile.getCanonicalPath();
            excelExporter.setFilePath( tmpFileName);
            excelExporter.output( book, bookdata, null);

            tmpFileName = excelExporter.getFilePath();
            tmpFile = new File( tmpFileName);
            converter.convert( tmpFile, new File( getFilePath()));

        } catch ( Exception e) {
            throw new ExportException( e);
        } finally {

            if ( tmpFile != null) {
                // EXCEL削除
                tmpFile.delete();
            }
            if ( !controlOfficeManager) {
                try {
                    officeManager.stop();
                } catch ( OfficeException e) {
                    throw new ExportException( e);
                }
            }
        }
    }

    /**
     * 変換フォーマット情報を作成する。
     *
     * @param configuration 変換情報
     * @return 変換フォーマット情報
     */
    private DocumentFormatRegistry createDocumentFormatRegistry( ConvertConfiguration configuration) {

        SimpleDocumentFormatRegistry registry = DefaultDocumentFormatRegistry.getInstance();

        if ( configuration == null || configuration.getOptionsProperties().isEmpty()) {
            return registry;
        }

        DocumentFormat documentFormat = registry.getFormatByExtension( "pdf");
        Map<String, Object> optionMap = new HashMap<String, Object>( documentFormat.getStoreProperties( DocumentFamily.SPREADSHEET));

        optionMap.put( "FilterData", configuration.getOptions());
        documentFormat.setStoreProperties( DocumentFamily.SPREADSHEET, optionMap);

        return registry;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.poireports.exporter.ReportBookExporter#getFormatType()
     */
    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.bbreak.excella.reports.exporter.ReportBookExporter#getExtention()
     */
    @Override
    public String getExtention() {
        return EXTENTION;
    }

}
