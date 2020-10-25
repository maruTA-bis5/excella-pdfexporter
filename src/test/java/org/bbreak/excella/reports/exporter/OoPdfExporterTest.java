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

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Workbook;
import org.jodconverter.office.ExternalOfficeManagerBuilder;
import org.jodconverter.office.OfficeManager;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.WorkbookTest;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;

/**
 * {@link org.bbreak.excella.reports.exporter.OoPdfExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class OoPdfExporterTest extends ReportsWorkbookTest {

    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    ConvertConfiguration configuration = null;

    private OfficeManager officeManager = new ExternalOfficeManagerBuilder().setPortNumber(8100).build();

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     * 
     * @throws IOException
     * @throws ExportException
     */
    @ParameterizedTest
    @CsvSource( WorkbookTest.VERSIONS)
    public void testOutput( String version) throws IOException, ExportException {

        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        String filePath = null;

        Workbook wb = getWorkbook( version);

        configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
        filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
        exporter.setFilePath( filePath);

        exporter.output( wb, new BookData(), configuration);
        File file = new File( exporter.getFilePath());
        assertTrue( file.exists());

        // オプション指定
        wb = getWorkbook( version);
        configuration.addOption( "PermissionPassword", "pass");
        configuration.addOption( "RestrictPermissions", Boolean.TRUE);
        configuration.addOption( "Printing", 0);
        configuration.addOption( "Changes", 4);
        filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
        exporter.setFilePath( filePath);

        exporter.output( wb, new BookData(), configuration);
        file = new File( exporter.getFilePath());
        assertTrue( file.exists());

        // 例外発生
        Workbook failingWb = getWorkbook( version);
        configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
        filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
        exporter.setFilePath( filePath);

        exporter.output( failingWb, new BookData(), configuration);

        file = new File( exporter.getFilePath());
        file.setReadOnly();
        assertThrows( ExportException.class, () -> exporter.output( failingWb, new BookData(), configuration));

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        assertEquals( "PDF", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        assertEquals( ".pdf", exporter.getExtention());
    }

}
