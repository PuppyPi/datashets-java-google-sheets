package rebound.richsheets.impls.live.googlesheets;

import static java.util.Collections.*;
import static rebound.testing.WidespreadTestingUtilities.*;
import static rebound.text.StringUtilities.*;
import static rebound.util.collections.CollectionUtilities.*;
import java.io.File;
import java.io.IOException;
import java.io.StringReader;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.List;
import javax.annotation.Nonnull;
import rebound.exceptions.ImPrettySureThisNeverActuallyHappensRuntimeException;
import rebound.exceptions.NotYetImplementedException;
import rebound.exceptions.UnexpectedHardcodedEnumValueException;
import rebound.richsheets.api.model.RichsheetsRow;
import rebound.richsheets.api.model.RichsheetsTable;
import rebound.richsheets.api.operation.RichsheetsConnection;
import rebound.richsheets.api.operation.RichsheetsOperation;
import rebound.richsheets.api.operation.RichsheetsOperation.RichsheetsOperationWithDataTimestamp;
import rebound.richsheets.api.operation.RichsheetsUnencodableFormatException;
import rebound.richsheets.api.operation.RichsheetsWriteData;
import rebound.richshets.model.cell.RichshetsCellContents;
import rebound.richshets.model.cell.RichshetsCellContents.RichshetsJustification;
import rebound.richshets.model.cell.RichshetsCellContents.RichshetsTextWrappingStrategy;
import rebound.richshets.model.cell.RichshetsCellContentsRun;
import rebound.richshets.model.cell.RichshetsCellContentsRun.RichshetsCellRunScriptLevel;
import rebound.richshets.model.cell.RichshetsColor;
import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.Sheets.Spreadsheets.Get;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.AppendDimensionRequest;
import com.google.api.services.sheets.v4.model.AutoResizeDimensionsRequest;
import com.google.api.services.sheets.v4.model.BatchUpdateSpreadsheetRequest;
import com.google.api.services.sheets.v4.model.CellData;
import com.google.api.services.sheets.v4.model.CellFormat;
import com.google.api.services.sheets.v4.model.Color;
import com.google.api.services.sheets.v4.model.DimensionProperties;
import com.google.api.services.sheets.v4.model.DimensionRange;
import com.google.api.services.sheets.v4.model.ExtendedValue;
import com.google.api.services.sheets.v4.model.GridCoordinate;
import com.google.api.services.sheets.v4.model.GridData;
import com.google.api.services.sheets.v4.model.GridProperties;
import com.google.api.services.sheets.v4.model.InsertDimensionRequest;
import com.google.api.services.sheets.v4.model.Request;
import com.google.api.services.sheets.v4.model.RowData;
import com.google.api.services.sheets.v4.model.Sheet;
import com.google.api.services.sheets.v4.model.SheetProperties;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.TextFormat;
import com.google.api.services.sheets.v4.model.TextFormatRun;
import com.google.api.services.sheets.v4.model.UpdateCellsRequest;
import com.google.api.services.sheets.v4.model.UpdateDimensionPropertiesRequest;
import com.google.api.services.sheets.v4.model.UpdateSheetPropertiesRequest;
import com.sun.istack.internal.Nullable;

//TODO warnings observer!

public class RichsheetsConnectionForGoogleSheets
{
	protected Credential auth;
	protected String applicationName;
	
	protected NetHttpTransport httpTransport;
	
	
	
	public RichsheetsConnectionForGoogleSheets(SimpleGoogleOAuthConfig auth) throws IOException, GeneralSecurityException
	{
		this(auth, "Richsheets");
	}
	
	public RichsheetsConnectionForGoogleSheets(SimpleGoogleOAuthConfig auth, String applicationName) throws IOException, GeneralSecurityException
	{
		httpTransport = GoogleNetHttpTransport.newTrustedTransport();
		this.auth = getCredentials(auth.getSecretsFile(), auth.getTokenCacheFile(), singletonList(SheetsScopes.SPREADSHEETS));
		this.applicationName = applicationName;
	}
	
	
	
	
	public RichsheetsConnection getConnectionFor(String spreadsheetId, int subsheetIndex)
	{
		return new RichsheetsConnection()
		{
			@Override
			public boolean isCapableOfAutoresizingColumns()
			{
				return true;
			}
			
			@Override
			public Date getCurrentLastModifiedTimestamp() throws IOException
			{
				return getLastModifiedTime(spreadsheetId);
			}
			
			@Override
			public void perform(Integer maxRowsToRead, RichsheetsOperation operation) throws IOException
			{
				Sheets service = new Sheets.Builder(httpTransport, JsonFactory, auth).setApplicationName(applicationName).build();
				
				
				
				refreshPinnedComments(spreadsheetId);
				
				
				
				//Read!
				Spreadsheet spreadsheet;
				{
					Get action = service.spreadsheets().get(spreadsheetId);
					action.setIncludeGridData(true);
					
					if (maxRowsToRead != null)
						action.setRanges(singletonList("1:"+maxRowsToRead));  //If there are less than N rows, this won't fail, it'll just use a least() function and return as many as it can (in this case of max=2, 0 or 1)  I tested this just now.  —Sean @ 2022-05-14 07:41:40 z
					
					spreadsheet = action.execute();
				}
				
				
				
				Date lastModifiedTimestampOfOriginalData;
				{
					lastModifiedTimestampOfOriginalData = getLastModifiedTime(spreadsheetId);
				}
				
				
				
				//Operate!
				final boolean readonly;
				final int columnsToAdd;
				final int rowsToAdd;
				final Integer setFrozenColumnsToThisOrDoNothingIfNull;
				final Integer setFrozenRowsToThisOrDoNothingIfNull;
				final Collection<Integer> columnsToAutoResize;
				final List<RowData> dataaaaaaaaaaaToWrite;
				final List<Integer> columnWidths;
				final List<Integer> rowHeights;
				{
					Sheet s = spreadsheet.getSheets().get(subsheetIndex);
					
					final int originalFrozenRowsCount;
					final int originalFrozenColumnsCount;
					{
						GridProperties g = s.getProperties().getGridProperties();
						Integer fr = g.getFrozenRowCount();
						Integer fc = g.getFrozenColumnCount();
						originalFrozenRowsCount = fr == null ? 0 : fr;
						originalFrozenColumnsCount = fc == null ? 0 : fc;
					}
					
					
					
					
					RichsheetsWriteData ro;
					
					//Do ittttttt!
					{
						if (operation != null)
						{
							RichsheetsTable ri = convertToRichsheets(s, originalFrozenColumnsCount, originalFrozenRowsCount);
							ro = operation instanceof RichsheetsOperationWithDataTimestamp ? ((RichsheetsOperationWithDataTimestamp)operation).performInMemory(ri, lastModifiedTimestampOfOriginalData) : operation.performInMemory(ri);
						}
						else
						{
							ro = null;
						}
					}
					
					
					if (ro == null)
					{
						readonly = true;
						columnsToAdd = 0;
						rowsToAdd = 0;
						setFrozenColumnsToThisOrDoNothingIfNull = null;
						setFrozenRowsToThisOrDoNothingIfNull = null;
						columnsToAutoResize = emptyList();
						dataaaaaaaaaaaToWrite = null;
						rowHeights = null;
						columnWidths = null;
					}
					else if (ro.getTable() == null)
					{
						readonly = false;
						columnsToAdd = 0;
						rowsToAdd = 0;
						setFrozenColumnsToThisOrDoNothingIfNull = null;
						setFrozenRowsToThisOrDoNothingIfNull = null;
						columnsToAutoResize = ro.getColumnsToAutoresize();
						dataaaaaaaaaaaToWrite = null;
						rowHeights = null;
						columnWidths = null;
					}
					else
					{
						readonly = false;
						columnsToAdd = ro.getTable().getNumberOfColumns() - getNumberOfColumns(s);
						rowsToAdd = ro.getTable().getNumberOfRows() - getNumberOfRows(s);
						setFrozenColumnsToThisOrDoNothingIfNull = ro.getTable().getFrozenColumns() == originalFrozenColumnsCount ? null : ro.getTable().getFrozenColumns();
						setFrozenRowsToThisOrDoNothingIfNull = ro.getTable().getFrozenRows() == originalFrozenRowsCount ? null : ro.getTable().getFrozenRows();
						columnsToAutoResize = ro.getColumnsToAutoresize();
						
						
						
						
						int newFrozenColumnsCount = ro.getTable().getFrozenColumns();
						int newFrozenRowsCount = ro.getTable().getFrozenRows();
						int newNumberOfColumns = ro.getTable().getNumberOfColumns();
						int newNumberOfRows = ro.getTable().getNumberOfRows();
						
						
						boolean[] booleanColumnsByNewIndex;  //columnIndexes in the intermediate form, not including frozen columns
						{
							if (maxRowsToRead != null)
							{
								booleanColumnsByNewIndex = null;
							}
							else
							{
								booleanColumnsByNewIndex = new boolean[newNumberOfColumns];
								
								for (int newColumnIndex = newFrozenColumnsCount; newColumnIndex < newNumberOfColumns; newColumnIndex++)
								{
									boolean booleanColumn;
									
									boolean atLeastOneActuallyBooleanable = false;
									for (int r = newFrozenRowsCount; r < newNumberOfRows; r++)
									{
										RichshetsCellContents ourCell = ro.getTable().getCell(newColumnIndex, r);
										String v = ourCell.justText();
										atLeastOneActuallyBooleanable |= v.equals("FALSE") || v.equals("TRUE");
									}
									
									if (atLeastOneActuallyBooleanable)
									{
										boolean atLeastOneDisqualifying = false;
										for (int r = newFrozenRowsCount; r < newNumberOfRows; r++)
										{
											RichshetsCellContents ourCell = ro.getTable().getCell(newColumnIndex, r);
											
											boolean isEmpty = ourCell.isEmptyText();
											
											if (!isEmpty)
											{
												String v = ourCell.justText();
												
												if (!v.equalsIgnoreCase("false") && !v.equalsIgnoreCase("true"))
												{
													atLeastOneDisqualifying = true;
													break;  //no point in going on!
												}
											}
										}
										
										//booleanColumn = atLeastOneActuallyBooleanable && !atLeastOneDisqualifying;
										booleanColumn = !atLeastOneDisqualifying;
									}
									else
									{
										booleanColumn = false;
									}
									
									
									booleanColumnsByNewIndex[newColumnIndex] = booleanColumn;
								}
							}
						}
						
						
						
						
						dataaaaaaaaaaaToWrite = new ArrayList<>();
						rowHeights = new ArrayList<>();
						columnWidths = ro.getTable().getColumnWidths();
						
						for (RichsheetsRow ourRow : ro.getTable().getRows())
						{
							RowData theirRow;
							{
								theirRow = new RowData();
								
								theirRow.setValues(mapToList(i -> encodeCell(ourRow.getCells().get(i), booleanColumnsByNewIndex == null ? false : booleanColumnsByNewIndex[i]), intervalIntegersList(0, newNumberOfColumns)));
							}
							
							dataaaaaaaaaaaToWrite.add(theirRow);
							rowHeights.add(ourRow.getHeight());
						}
					}
				}
				
				
				
				
				
				
				
				
				
				
				
				
				//Write!
				if (!readonly)
				{
					// https://googleapis.dev/java/google-api-services-sheets/latest/com/google/api/services/sheets/v4/model/Request.html
					// https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request
					List<Request> reqs;
					{
						reqs = new ArrayList<>();
						
						//Expanding the sheet must come before everything else!
						{
							if (columnsToAdd > 0)
							{
								AppendDimensionRequest req = new AppendDimensionRequest();
								req.setDimension("COLUMNS");
								req.setSheetId(subsheetIndex);
								req.setLength(rowsToAdd);
								
								Request r = new Request();  //note that these can only hold one typed request!
								r.setAppendDimension(req);
								reqs.add(r);
							}
							
							
							if (rowsToAdd > 0)
							{
								if (maxRowsToRead == null)
								{
									AppendDimensionRequest req = new AppendDimensionRequest();
									req.setDimension("ROWS");
									req.setSheetId(subsheetIndex);
									req.setLength(rowsToAdd);
									
									Request r = new Request();  //note that these can only hold one typed request!
									r.setAppendDimension(req);
									reqs.add(r);
								}
								else
								{
									InsertDimensionRequest req = new InsertDimensionRequest();
									
									DimensionRange range = new DimensionRange();
									range.setDimension("ROWS");
									range.setSheetId(subsheetIndex);
									range.setStartIndex(maxRowsToRead);
									range.setEndIndex(maxRowsToRead + rowsToAdd);
									
									req.setRange(range);
									
									Request r = new Request();  //note that these can only hold one typed request!
									r.setInsertDimension(req);
									reqs.add(r);
								}
							}
						}
						
						
						
						//Set frozen columns/rows
						if (setFrozenColumnsToThisOrDoNothingIfNull != null || setFrozenRowsToThisOrDoNothingIfNull != null)
						{
							GridProperties gp = new GridProperties();
							
							if (setFrozenColumnsToThisOrDoNothingIfNull != null)
								gp.setFrozenColumnCount(setFrozenColumnsToThisOrDoNothingIfNull);
							
							if (setFrozenRowsToThisOrDoNothingIfNull != null)
								gp.setFrozenRowCount(setFrozenRowsToThisOrDoNothingIfNull);
							
							SheetProperties props = new SheetProperties();
							props.setGridProperties(gp);
							
							UpdateSheetPropertiesRequest req = new UpdateSheetPropertiesRequest();
							req.setProperties(props);
							
							Request r = new Request();  //note that these can only hold one typed request!
							r.setUpdateSheetProperties(req);
							reqs.add(r);
						}
						
						
						
						
						
						//Set! The! Dataaaaaaaaaaaa!  \:D/
						if (dataaaaaaaaaaaToWrite != null)
						{
							GridCoordinate origin = new GridCoordinate();
							origin.setColumnIndex(0);
							origin.setRowIndex(0);
							
							UpdateCellsRequest req = new UpdateCellsRequest();
							req.setStart(origin);
							req.setRows(dataaaaaaaaaaaToWrite);
							
							Request r = new Request();  //note that these can only hold one typed request!
							r.setUpdateCells(req);
							reqs.add(r);
						}
						
						
						
						
						//Set the row heights!
						if (rowHeights != null)
						{
							int n = rowHeights.size();
							
							for (int rowIndex = 0; rowIndex < n; rowIndex++)  //Todo find contiguous ranges for a minor optimization ^^'
							{
								Integer h = rowHeights.get(rowIndex);
								
								DimensionRange range = new DimensionRange();
								range.setSheetId(subsheetIndex);
								range.setDimension("ROWS");
								range.setStartIndex(rowIndex);  //inclusive
								range.setEndIndex(rowIndex+1);  //exclusive
								
								DimensionProperties props = new DimensionProperties();
								props.setPixelSize(h == null ? DefaultGoogleSheetsRowHeight : h);
								
								UpdateDimensionPropertiesRequest req = new UpdateDimensionPropertiesRequest();
								req.setProperties(props);
								req.setRange(range);
								
								Request r = new Request();  //note that these can only hold one typed request!
								r.setUpdateDimensionProperties(req);
								reqs.add(r);
							}
						}
						
						
						
						
						
						//Set the column widths!
						if (columnWidths != null)
						{
							int n = columnWidths.size();
							
							for (int columnIndex = 0; columnIndex < n; columnIndex++)  //Todo find contiguous ranges for a minor optimization ^^'
							{
								Integer h = columnWidths.get(columnIndex);
								
								DimensionRange range = new DimensionRange();
								range.setSheetId(subsheetIndex);
								range.setDimension("COLUMNS");
								range.setStartIndex(columnIndex);  //inclusive
								range.setEndIndex(columnIndex+1);  //exclusive
								
								DimensionProperties props = new DimensionProperties();
								props.setPixelSize(h == null ? DefaultGoogleSheetsColumnWidth : h);
								
								UpdateDimensionPropertiesRequest req = new UpdateDimensionPropertiesRequest();
								req.setProperties(props);
								req.setRange(range);
								
								Request r = new Request();  //note that these can only hold one typed request!
								r.setUpdateDimensionProperties(req);
								reqs.add(r);
							}
						}
						
						
						
						
						//Resizing columns (afterrrrrrr setting data and column widths if we do that! :D )
						for (int columnIndex : columnsToAutoResize)  //Todo find contiguous ranges for a minor optimization ^^'
						{
							DimensionRange dims = new DimensionRange();
							dims.setSheetId(subsheetIndex);
							dims.setDimension("COLUMNS");
							dims.setStartIndex(columnIndex);  //inclusive
							dims.setEndIndex(columnIndex+1);  //exclusive
							
							AutoResizeDimensionsRequest req = new AutoResizeDimensionsRequest();
							req.setDimensions(dims);
							
							Request r = new Request();  //note that these can only hold one typed request!
							r.setAutoResizeDimensions(req);
							reqs.add(r);
						}
					}
					
					
					BatchUpdateSpreadsheetRequest mainreq = new BatchUpdateSpreadsheetRequest();
					mainreq.setIncludeSpreadsheetInResponse(false);
					mainreq.setResponseIncludeGridData(false);
					mainreq.setRequests(reqs);
					
					service.spreadsheets().batchUpdate(spreadsheetId, mainreq).execute();
				}
			}
		};
	}
	
	
	
	
	protected RichsheetsTable convertToRichsheets(Sheet s, int frozenColumnsCount, int frozenRowsCount)
	{
		List<GridData> gds = s.getData();
		
		if (gds.size() != 1)
			throw new RuntimeException("What does it meeeeeeean to have multiple GridData's?!");  //Todo figure out what it means XD   (it's not more than one even when there are multiple (sub)sheets in my testing!  —Sean @ 2022-05-14 07 z)
		GridData gd = gds.get(0);
		
		
		int numberOfColumns = gd.getColumnMetadata().size();
		int numberOfRows = gd.getRowMetadata().size();
		
		
		final List<RowData> googleSheetsRows;
		{
			final List<RowData> od = gd.getRowData();
			
			//Google Sheets returns null here not an empty list if the spreadsheet is empty! (even if there are frozen rows)  I tested this just now.   —Sean @ 2022-05-14 07:43:07 z
			googleSheetsRows = od == null ? emptyList() : od;
		}
		
		
		asrt(numberOfRows == googleSheetsRows.size());
		
		
		
		List<Integer> columnWidths = mapToList(columnIndex -> gd.getColumnMetadata().get(columnIndex).getPixelSize(), intervalIntegersList(0, numberOfColumns));
		
		
		List<RichsheetsRow> rows = mapToList(rowIndex ->
		{
			List<RichshetsCellContents> cells = new ArrayList<>(numberOfColumns);
			
			for (CellData gsc : googleSheetsRows.get(rowIndex).getValues())
				cells.add(decodeCell(gsc));
			
			for (int i = cells.size(); i < numberOfColumns; i++)
				cells.add(RichshetsCellContents.Blank);
			
			Integer rowHeight = gd.getRowMetadata().get(rowIndex).getPixelSize();
			
			return new RichsheetsRow(cells, rowHeight);
			
		}, intervalIntegersList(0, numberOfRows));
		
		
		
		RichsheetsTable rt = new RichsheetsTable(rows);
		rt.setColumnWidths(columnWidths);
		rt.setFrozenColumns(frozenColumnsCount);
		rt.setFrozenRows(frozenRowsCount);
		return rt;
	}
	
	
	
	
	
	
	
	protected static int getNumberOfColumns(Sheet s)
	{
		List<GridData> gds = s.getData();
		
		if (gds.size() != 1)
			throw new RuntimeException("What does it meeeeeeean to have multiple GridData's?!");  //Todo figure out what it means XD   (it's not more than one even when there are multiple (sub)sheets in my testing!  —Sean @ 2022-05-14 07 z)
		GridData gd = gds.get(0);
		
		
		return gd.getColumnMetadata().size();
	}
	
	
	protected static int getNumberOfRows(Sheet s)
	{
		List<GridData> gds = s.getData();
		
		if (gds.size() != 1)
			throw new RuntimeException("What does it meeeeeeean to have multiple GridData's?!");  //Todo figure out what it means XD   (it's not more than one even when there are multiple (sub)sheets in my testing!  —Sean @ 2022-05-14 07 z)
		GridData gd = gds.get(0);
		
		
		return gd.getRowMetadata().size();
	}
	
	
	
	
	
	
	
	//	protected static boolean isSheetsCellBoolean(CellData sheetsCell)
	//	{
	//		return sheetsCell.getEffectiveValue().getBoolValue() != null;
	//	}
	
	
	/**
	 * @return null if column is out of bounds (since Google Sheets API can return different-sized rows if there's blank cells on the end of one), instead of throwing a {@link IndexOutOfBoundsException}!
	 * @throws IndexOutOfBoundsException if the row index is out of bounds.
	 */
	protected static @Nullable CellData getSheetsCell(List<RowData> rows, int columnIndex, int rowIndex) throws IndexOutOfBoundsException
	{
		List<CellData> cells = rows.get(rowIndex).getValues();
		
		if (columnIndex >= cells.size())
			return null;
		else
			return cells.get(columnIndex);
	}
	
	
	protected static RichshetsCellContents decodeCell(CellData gsCell)
	{
		String v = gsCell.getFormattedValue();
		
		CellFormat f = gsCell.getEffectiveFormat();
		
		List<TextFormatRun> gsruns = gsCell.getTextFormatRuns();
		
		int n = gsruns.size();
		
		List<RichshetsCellContentsRun> rsruns = mapToList(i ->
		{
			TextFormatRun r = gsruns.get(i);
			
			TextFormat tf = r.getFormat();
			int start = r.getStartIndex();
			
			int end = i == n - 1 ? v.length() : gsruns.get(i+1).getStartIndex();
			
			Boolean bold = tf.getBold();
			Boolean italic = tf.getItalic();
			Boolean underline = tf.getUnderline();
			Boolean strikethrough = tf.getStrikethrough();
			Color fgcol = tf.getForegroundColor();
			
			String t = v.substring(start, end);
			
			//Google Sheets doesn't support superscript or subscript
			
			return new RichshetsCellContentsRun(t, fin(bold), fin(underline), fin(italic), fin(strikethrough), RichshetsCellRunScriptLevel.Normal, decodeColor(fgcol));
			
		}, intervalIntegersList(0, n));
		
		
		
		RichshetsJustification justification;
		{
			String s = f.getHorizontalAlignment();
			
			if (s == null)
				justification = null;
			else if ("LEFT".equals(s))
				justification = RichshetsJustification.Left;
			else if ("CENTER".equals(s))
				justification = RichshetsJustification.Center;
			else if ("RIGHT".equals(s))
				justification = RichshetsJustification.Right;
			else if ("HORIZONTAL_ALIGN_UNSPECIFIED".equals(s))  //does this ever actually get returned?
				justification = null;
			else
				throw new ImPrettySureThisNeverActuallyHappensRuntimeException("Google Sheets Horizontal Alignment: "+repr(s));
		}
		
		Color bgcol = f.getBackgroundColor();
		
		RichshetsTextWrappingStrategy wrap;
		{
			String s = f.getWrapStrategy();
			
			if ("WRAP".equals(s))
				wrap = RichshetsTextWrappingStrategy.Wrap;
			else if ("OVERFLOW".equals(s))
				wrap = RichshetsTextWrappingStrategy.Overflow;
			else if ("CLIP".equals(s))
				wrap = RichshetsTextWrappingStrategy.Clip;
			else
				throw new ImPrettySureThisNeverActuallyHappensRuntimeException("Google Sheets Wrapping Strategy: "+repr(s));
		}
		
		return new RichshetsCellContents(rsruns, justification, decodeColor(bgcol), wrap);
	}
	
	
	
	
	protected static CellData encodeCell(RichshetsCellContents datashetsCell, boolean bool)
	{
		List<TextFormatRun> gsruns;
		{
			gsruns = new ArrayList<>();
			
			int i = 0;
			for (RichshetsCellContentsRun rsrun : datashetsCell.getContents())
			{
				TextFormat tf;
				{
					tf = new TextFormat();
					
					tf.setBold(rsrun.isBold());
					tf.setItalic(rsrun.isItalic());
					tf.setUnderline(rsrun.isUnderline());
					tf.setStrikethrough(rsrun.isStrikethrough());
					
					if (rsrun.getScriptLevel() != RichshetsCellRunScriptLevel.Normal)
						throw new RichsheetsUnencodableFormatException("Google Sheets cannot encode superscript or subscript rich text.");
					
					tf.setForegroundColor(encodeColor(rsrun.getTextColor()));
				}
				
				
				TextFormatRun gsrun = new TextFormatRun();
				gsrun.setStartIndex(i);
				gsrun.setFormat(tf);
				gsruns.add(gsrun);
				
				i += rsrun.getContents().length();
			}
		}
		
		
		
		CellFormat f = new CellFormat();
		f.setBackgroundColor(encodeColor(datashetsCell.getBackgroundColor()));
		
		
		//Justification
		{
			RichshetsJustification j = datashetsCell.getJustification();
			if (j == RichshetsJustification.Left)
				f.setHorizontalAlignment("LEFT");
			else if (j == RichshetsJustification.Center)
				f.setHorizontalAlignment("CENTER");
			else if (j == RichshetsJustification.Right)
				f.setHorizontalAlignment("RIGHT");
			else
				throw new UnexpectedHardcodedEnumValueException(j);
		}
		
		
		//Wrap strategy
		{
			RichshetsTextWrappingStrategy w = datashetsCell.getWrappingStrategy();
			if (w == RichshetsTextWrappingStrategy.Wrap)
				f.setHorizontalAlignment("WRAP");
			else if (w == RichshetsTextWrappingStrategy.Overflow)
				f.setHorizontalAlignment("OVERFLOW");
			else if (w == RichshetsTextWrappingStrategy.Clip)
				f.setHorizontalAlignment("CLIP");
			else
				throw new UnexpectedHardcodedEnumValueException(w);
		}
		
		
		CellData sheetsCell = new CellData();
		sheetsCell.setEffectiveFormat(f);
		
		String text = datashetsCell.justText();
		
		if (bool)
		{
			boolean value;
			{
				if (text.equalsIgnoreCase("true"))
					value = true;
				else if (text.equalsIgnoreCase("false"))
					value = false;
				else
					throw new IllegalArgumentException("Not a boolean value!: '"+text+"'");
			}
			
			ExtendedValue ev = new ExtendedValue();
			ev.setBoolValue(value);
			
			sheetsCell.setUserEnteredValue(ev);
			sheetsCell.setEffectiveValue(ev);
			sheetsCell.setFormattedValue(value ? "TRUE" : "FALSE");
		}
		else
		{
			ExtendedValue ev = new ExtendedValue();
			ev.setStringValue(text);
			
			sheetsCell.setUserEnteredValue(ev);
			sheetsCell.setEffectiveValue(ev);
			sheetsCell.setFormattedValue(text);
		}
		
		return sheetsCell;
	}
	
	
	
	
	
	protected static @Nullable RichshetsColor decodeColor(@Nonnull Color sheetsColor)
	{
		Float r = sheetsColor.getRed();
		Float g = sheetsColor.getGreen();
		Float b = sheetsColor.getBlue();
		
		if (r == null && g == null && b == null)
			return null;
		else
			return new RichshetsColor(cf2i(r == null ? 0 : r), cf2i(g == null ? 0 : g), cf2i(b == null ? 0 : b));
	}
	
	protected static @Nonnull Color encodeColor(@Nullable RichshetsColor datashetsColor)
	{
		Color c = new Color();
		if (datashetsColor != null)
		{
			c.setRed(ci2f(datashetsColor.getR()));
			c.setGreen(ci2f(datashetsColor.getG()));
			c.setBlue(ci2f(datashetsColor.getB()));
		}
		return c;
	}
	
	protected static int cf2i(float f)
	{
		if (f > 1)
			return 255;
		else if (f < 0)
			return 0;
		else
			return Math.round(f * 255);
	}
	
	protected static float ci2f(int i)
	{
		if (i < 0)  throw new IllegalArgumentException();
		if (i > 255)  throw new IllegalArgumentException();
		
		return i / 255f;
	}
	
	protected static boolean fin(Boolean t)
	{
		return t == null ? false : t;
	}
	
	
	
	
	protected static RowData newRowData(List<CellData> l)
	{
		RowData r = new RowData();
		r.setValues(l);
		return r;
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	/**
	 * Make sure to call this explicitly as it's not part of normal maintenance!
	 */
	public void refreshPinnedComments(String spreadsheetId)
	{
		// https://developers.google.com/drive/api/guides/manage-comments
		//TODO Google Drive api! :>
	}
	
	
	public Date getLastModifiedTime(String spreadsheetId)
	{
		//TODO Google Drive api! :>
		return null;
	}
	
	
	
	
	/**
	 * This is useful for figuring out if a table is huge before running {@link RichsheetsConnection#perform(Integer, RichsheetsOperation)} or similar on it (namely if you might want to do a HEAD style validation/maintenance if it's very large)
	 * @return {width (number of columns), height (number of rows)} in the Google Sheet  (not the "Datashet" which cares about header rows vs. data rows and such, just the raw underlying Google Sheet X3   ..we can make another function for that if we want it :3 )
	 */
	public int[] getDimensions(String spreadsheetId)
	{
		//TODO get the size without actually downloading allllllllll the dataaaaaaaaa!  XD''
		throw new NotYetImplementedException();
	}
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	/**
	 * @param scopes Eg, {@link SheetsScopes#SPREADSHEETS_READONLY}, {@link SheetsScopes#SPREADSHEETS}
	 */
	protected static Credential getCredentials(File secretsFile, File tokenCacheFile, Collection<String> scopes) throws IOException, GeneralSecurityException
	{
		final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
		
		// Load client secrets.
		String t = ReboundStaticallyCopiedUtilities.readAllText(secretsFile);
		
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JsonFactory, new StringReader(t));
		
		// Build flow and trigger user authorization request.
		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(HTTP_TRANSPORT, JsonFactory, clientSecrets, scopes).setDataStoreFactory(new FileDataStoreFactory(tokenCacheFile)).setAccessType("offline").build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}
	
	
	protected static final JsonFactory JsonFactory = JacksonFactory.getDefaultInstance();
	
	protected static final int DefaultGoogleSheetsColumnWidth = 100;
	protected static final int DefaultGoogleSheetsRowHeight = 21;
}
