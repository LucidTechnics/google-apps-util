package com.lucidtechnics.util;

import com.google.gdata.client.DocumentQuery;
import com.google.gdata.client.docs.DocsService;
import com.google.gdata.client.spreadsheet.CellQuery;
import com.google.gdata.client.spreadsheet.ListQuery;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.Link;
import com.google.gdata.data.MediaContent;
import com.google.gdata.data.PlainTextConstruct;
import com.google.gdata.data.acl.AclEntry;
import com.google.gdata.data.acl.AclRole;
import com.google.gdata.data.acl.AclScope;
import com.google.gdata.data.batch.BatchOperationType;
import com.google.gdata.data.batch.BatchStatus;
import com.google.gdata.data.batch.BatchUtils;
import com.google.gdata.data.docs.DocumentListEntry;
import com.google.gdata.data.docs.DocumentListFeed;
import com.google.gdata.data.docs.FolderEntry;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.ListEntry;
import com.google.gdata.data.spreadsheet.ListFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;

public class GoogleDocsUtility 
{

    private final static Logger logger = Logger.getLogger("GoogleDocsUtility");

    /**
     * Our view of Google Spreadsheets as an authenticated Google user.
     */
    private SpreadsheetService spreadsheetService;

    private DocsService docsService;

    private final static String APPLICATION_NAME = "Foo";

    /**
     * Log in to Google, under the Google Spreadsheets and Docs account.
     *
     * @param username name of user to authenticate (e.g. yourname@gmail.com)
     * @param password password to use for authentication
     * @throws AuthenticationException if the service is unable to validate the
     *                                 username and password.
     */
    public void login(String username, String password)
            throws AuthenticationException
    {
        spreadsheetService = new SpreadsheetService(APPLICATION_NAME);
        spreadsheetService.setUserCredentials(username, password);
        spreadsheetService.setConnectTimeout(0);
        spreadsheetService.setReadTimeout(0);

        docsService = new DocsService(APPLICATION_NAME);
        docsService.setUserCredentials(username, password);
        docsService.setConnectTimeout(0);
        docsService.setReadTimeout(0);

    }

    public ListFeed query(ListQuery query) throws IOException, ServiceException
    {
        logger.info("Query: " +query.getSpreadsheetQuery());

        return spreadsheetService.query(query, ListFeed.class);
    }

	public CellFeed queryCells(CellQuery query) throws IOException, ServiceException
	{
		logger.info("\n\nqueryCells: "+ query);
//		CellFeed feed = spreadsheetService.query(query, CellFeed.class);

		// for (CellEntry cell : feed.getEntries())
		// {
		//   logger.info(cell.getTitle().getPlainText());
		//   String shortId = cell.getId().substring(cell.getId().lastIndexOf('/') + 1);
		//   // logger.info(" -- Cell(" + shortId + "/" + cell.getTitle().getPlainText()
		//   //     + ") formula(" + cell.getCell().getInputValue() + ") numeric("
		//   //     + cell.getCell().getNumericValue() + ") value("
		//   //     + cell.getCell().getValue() + ")");        
		// }
		
		return spreadsheetService.query(query, CellFeed.class);
	}

    public void copyRowsFromCellFeed(CellFeed _cellFeed, WorksheetEntry _targetWorksheet) throws IOException, ServiceException
	{
				
	   HashMap m = new HashMap();
	   List<Map<Integer, Object>> rows = new ArrayList<Map<Integer, Object>>();

		int currentRow = 1;

		for (CellEntry cell : _cellFeed.getEntries()) 
		{
			if(currentRow != cell.getCell().getRow())
			{
				rows.add(m);
				currentRow=cell.getCell().getRow();
				m = new HashMap();
			}
			m.put(cell.getCell().getCol(), cell.getCell().getInputValue());

//			 logger.info(cell.getTitle().getPlainText());
//			  String shortId = cell.getId().substring(cell.getId().lastIndexOf('/') + 1);
//			  logger.info(" -- A Cell(" + shortId + "/" + cell.getTitle().getPlainText()
//			      + ") formula(" + cell.getCell().getInputValue() + ") numeric("
//			      + cell.getCell().getNumericValue() + ") value("
//			      + cell.getCell().getValue() + ")");
		}
		

		//add last row
		rows.add(m);
		logger.info("\n\nRows to add: "+rows.size());
		
		writeBatchRows(_targetWorksheet, rows, 1);
		
	}
	
	
	public List<WorksheetEntry> copyDocument(DocumentListEntry documentListEntry, String newDocumentTitle, DocumentListEntry destinationFolderEntry)
			throws IOException, ServiceException
    {

        documentListEntry.setTitle(new PlainTextConstruct(newDocumentTitle));

        URL url = new URL("https://docs.google.com/feeds/default/private/full/");

        DocumentListEntry newDoc = docsService.insert(url, documentListEntry);

        String destFolder = ((MediaContent) destinationFolderEntry.getContent()).getUri();
        URL newUrl = new URL (destFolder);

        //Convert DocumentListEntry to SpreadsheetEntry
        DocumentListEntry newDocumentListEntry = docsService.insert(newUrl, newDoc);
        String spreadsheetURL = "https://spreadsheets.google.com/feeds/spreadsheets/" + newDocumentListEntry.getDocId();

        SpreadsheetEntry spreadsheetEntry = spreadsheetService.getEntry(new URL(spreadsheetURL), SpreadsheetEntry.class);

        return spreadsheetEntry.getWorksheets();
    }

	public List<WorksheetEntry> createDocument(String _newDocumentTitle, DocumentListEntry _destinationFolderEntry)
			throws IOException, ServiceException
	{
		DocumentListEntry newEntry = new com.google.gdata.data.docs.SpreadsheetEntry();
		newEntry.setTitle(new PlainTextConstruct(_newDocumentTitle));
		
		URL url = new URL("https://docs.google.com/feeds/default/private/full/");

		DocumentListEntry newDoc = docsService.insert(url, newEntry);

		String destFolder = ((MediaContent) _destinationFolderEntry.getContent()).getUri();
		URL newUrl = new URL (destFolder);

		// Convert DocumentListEntry to SpreadsheetEntry
		DocumentListEntry newDocumentListEntry = docsService.insert(newUrl, newDoc);
		String spreadsheetURL = "https://spreadsheets.google.com/feeds/spreadsheets/" + newDocumentListEntry.getDocId();

		SpreadsheetEntry spreadsheetEntry = spreadsheetService.getEntry(new URL(spreadsheetURL), SpreadsheetEntry.class);

		return spreadsheetEntry.getWorksheets();
	}

	public SpreadsheetEntry getSpreadsheetEntryFromDocumentListEntry(DocumentListEntry docListEntry) throws IOException, ServiceException
	{
		
        String spreadsheetURL = "https://spreadsheets.google.com/feeds/spreadsheets/" + docListEntry.getDocId();

        return spreadsheetService.getEntry(new URL(spreadsheetURL), SpreadsheetEntry.class);
	}

    public DocumentListEntry getDocument(URL folderFeedUri, String documentName) throws IOException, ServiceException
    {
		DocumentQuery query = new DocumentQuery(folderFeedUri);
		query.setTitleQuery(documentName);
		query.setTitleExact(true);
		query.setMaxResults(1);
		DocumentListFeed feed = docsService.getFeed(query, DocumentListFeed.class);
        DocumentListEntry doc = null;
        if(! feed.getEntries().isEmpty())
        {
            doc = feed.getEntries().get(0);
        }

        return doc;
	}

    public DocumentListEntry createFolder(String folderName, String roleString, List<String> emails) throws IOException, ServiceException
    {
        logger.info("\n\nFolder Exists? " + folderName );
        boolean folderFound = false;

        URL feedUri = new URL("https://docs.google.com/feeds/default/private/full/-/folder");
        DocumentListFeed feed = docsService.getFeed(feedUri, DocumentListFeed.class);
        DocumentListEntry folderEntry = new DocumentListEntry();

        for (DocumentListEntry entry : feed.getEntries())
        {
            if(folderName.equalsIgnoreCase(entry.getTitle().getPlainText()))
            {
                folderFound = true;
                folderEntry = entry;
                break;
            }
        }

        if(!folderFound)
        {
            logger.info("\n\nCreating folder " + folderName);
            DocumentListEntry newEntry = new FolderEntry();
            newEntry.setTitle(new PlainTextConstruct(folderName));
            URL feedUrl = new URL("https://docs.google.com/feeds/default/private/full/");
            folderEntry = docsService.insert(feedUrl, newEntry);

            shareResource(roleString, emails, folderEntry);

        }

        return folderEntry;
    }
 
    public String getFolderURI(String _folderName) throws IOException, ServiceException
    {

        URL feedUri = new URL("https://docs.google.com/feeds/default/private/full/-/folder");
        DocumentListFeed feed = docsService.getFeed(feedUri, DocumentListFeed.class);
        DocumentListEntry folderEntry = new DocumentListEntry();

        for (DocumentListEntry entry : feed.getEntries())
        {
            if(_folderName.equalsIgnoreCase(entry.getTitle().getPlainText()))
            {
                folderEntry = entry;
				break;
            }
        }
		String folderURIString = null;

		if(folderEntry.getContent() != null)
		{
			folderURIString = getDocumentFeedUri(folderEntry);
		}
		
		
		return folderURIString;
    }

	 public DocumentListFeed getDocsFeedForFolder(String _folderResourceId) throws IOException, ServiceException
	 {
		URL url = new URL(_folderResourceId);
		logger.info("folder url: " + url);
	    return docsService.getFeed(url, DocumentListFeed.class);
	  }


    public void shareResource(String roleString, String email, DocumentListEntry documentListEntry) throws IOException
    {
        AclRole role = new AclRole(roleString);
        AclScope scope = new AclScope(AclScope.Type.USER, email);

        AclEntry aclEntry = new AclEntry();
        aclEntry.setRole(role);
        aclEntry.setScope(scope);

		try
		{
        	docsService.insert(new URL(documentListEntry.getAclFeedLink().getHref()), aclEntry);
		} catch (ServiceException se)
		{
			logger.info("Unable to share resource or resource already shared.");
		}
    }

    public void shareResource(String roleString, List<String> emails, DocumentListEntry documentListEntry) throws IOException
    {
       	for(String email : emails)
	   	{
			shareResource(roleString, email, documentListEntry);
		}
    }


    public String getTagValue(ListFeed listFeed, String tagQuery)
    {

        String tagValue = "";

        for (ListEntry entry : listFeed.getEntries())
        {
            for (String tag : entry.getCustomElements().getTags())
            {
                if(tag.equalsIgnoreCase(tagQuery))
                {
                    tagValue = entry.getCustomElements().getValue(tag);
                    break;
                }
            }
        }
        return tagValue;
    }

    public URL createWorksheet(URL url, String worksheetTitle, int rows, int columns, List<String> columnNames) throws IOException, ServiceException
    {
        WorksheetEntry worksheet = new WorksheetEntry();
        worksheet.setTitle(new PlainTextConstruct(worksheetTitle));

        worksheet.setRowCount(rows);
        worksheet.setColCount(columns);

        WorksheetEntry createdWorksheet = spreadsheetService.insert(url, worksheet);

        //create header
        CellFeed cellFeed = spreadsheetService.getFeed (createdWorksheet.getCellFeedUrl(), CellFeed.class);
        int i = 1;

        for(String columnName : columnNames)
        {
            CellEntry cellEntry = new CellEntry (1, i, columnName);
            cellFeed.insert (cellEntry);
            i++;
        }
        return createdWorksheet.getListFeedUrl();
    }

    public WorksheetEntry createWorksheet(URL url, String worksheetTitle, int rows, int columns) throws IOException, ServiceException
    {
        WorksheetEntry worksheet = new WorksheetEntry();
        worksheet.setTitle(new PlainTextConstruct(worksheetTitle));

        worksheet.setRowCount(rows);
        worksheet.setColCount(columns);

        return spreadsheetService.insert(url, worksheet);
    }

    public WorksheetEntry createWorksheet(URL _worksheetUrl, WorksheetEntry _worksheetEntry) throws IOException, ServiceException
    {

        WorksheetEntry worksheet = new WorksheetEntry(_worksheetEntry);

        return spreadsheetService.insert(_worksheetUrl, worksheet);
    }


    public void addEmptyRows(WorksheetEntry worksheetEntry, int newRows) throws IOException, ServiceException
    {
        int rows = worksheetEntry.getRowCount();
        worksheetEntry.setRowCount(rows+newRows);
        worksheetEntry.update();
    }

    /**
    * Logs a list of the entries in the batch request in a human
    * readable format.
    *
    * @param batchRequest the CellFeed containing entries to display.
    */
    private void printBatchRequest(CellFeed batchRequest)
    {
        logger.info("Current operations in batch");
        for (CellEntry entry : batchRequest.getEntries()) {
          String msg = "\tID: " + BatchUtils.getBatchId(entry) + " - "
              + BatchUtils.getBatchOperationType(entry) + " row: "
              + entry.getCell().getRow() + " col: " + entry.getCell().getCol()
              + " value: " + entry.getCell().getInputValue();
          logger.info(msg);
        }
    }

    public void writeBatchRows(WorksheetEntry worksheetEntry, List<Map<Integer, Object>> rowList, int rowOffset) throws IOException, ServiceException
    {

       long startTime = System.currentTimeMillis();

        URL cellFeedUrl = worksheetEntry.getCellFeedUrl();
        CellFeed cellFeed = spreadsheetService.getFeed(worksheetEntry.getCellFeedUrl(), CellFeed.class);
        CellFeed batchRequest = new CellFeed();
        logger.info("Get Row Count:  " + cellFeed.getRowCount());
        int rowToBegin = rowOffset;
        addEmptyRows(worksheetEntry,rowList.size());

        logger.info("Row To Begin:  " + rowToBegin);

        // Build list of cell addresses to be filled in
        List<CellAddress> cellAddrs = new ArrayList<CellAddress>();
        CellAddress cellAddress = new CellAddress();
        String formula;
        for (Map<Integer,Object> row : rowList)
        {

            for (Map.Entry<Integer, Object> entry : row.entrySet())
            {
                int column = entry.getKey();
				if(!(entry.getValue() instanceof String))
				{
					formula = entry.getValue().toString();
				} else 
				{
                	formula = (String) entry.getValue();
				}
                logger.info("********************Column: " + column + "Formula: " + formula);

                cellAddress.setCol(column);
                cellAddress.setRow(rowToBegin);
                cellAddress.setIdString(String.format("R%sC%s", rowToBegin, column));
                cellAddrs.add(cellAddress);
                for (CellAddress cellAddr : cellAddrs)
                {
                    CellEntry batchEntry = new CellEntry(cellAddr.row, cellAddr.col, formula);
                    batchEntry.setId(String.format("%s/%s", cellFeedUrl.toString(), cellAddr.idString));
                    BatchUtils.setBatchId(batchEntry, cellAddr.idString);
                    BatchUtils.setBatchOperationType(batchEntry, BatchOperationType.UPDATE);

                    logger.fine("Batch Entry: " + batchEntry);

                    batchRequest.getEntries().add(batchEntry);
                    batchEntry = null;
                }
                cellAddrs.clear();

            }
            // batch per row
            if(rowToBegin % 100 == 0)
            {
                long startBatchTime = System.currentTimeMillis();

                logger.info("\n\n\nEvery 100 rows batch call: " + rowToBegin);
                performBatchUpdate(batchRequest, cellFeedUrl);
                batchRequest.getEntries().clear();
                logger.info("\n\n ms elapsed for batch:  " + (System.currentTimeMillis() - startBatchTime));
            }

            rowToBegin++;

        }

		logger.info("\n\n\n\nms elapsed to create batch request: " + (System.currentTimeMillis() - startTime));
        //for the stragglers
        logger.info("\n\n\nLast rows batch call: " + rowToBegin);
        performBatchUpdate(batchRequest, cellFeedUrl);

	}

    public void performBatchUpdate(CellFeed batchRequest, URL cellFeedUrl) throws IOException, ServiceException
    {

        logger.info("\n Begin Batch Update \n");

       long startTime = System.currentTimeMillis();

        // printBatchRequest(batchRequest);

        CellFeed cellFeed = spreadsheetService.getFeed(cellFeedUrl, CellFeed.class);

        // Submit the update
        //set header work around for http://code.google.com/p/gdata-java-client/issues/detail?id=103
        spreadsheetService.setHeader("If-Match", "*");
        Link batchLink = cellFeed.getLink(Link.Rel.FEED_BATCH, Link.Type.ATOM);
//        CellFeed batchResponse = spreadsheetService.batch(new URL(batchLink.getHref()), batchRequest);
        spreadsheetService.batch(new URL(batchLink.getHref()), batchRequest);
        spreadsheetService.setHeader("If-Match", null);

		logger.info("\n ms elapsed for batch update: \n" + (System.currentTimeMillis() - startTime));

       // Check the results
//        boolean isSuccess = true;
//        for (CellEntry entry : batchResponse.getEntries()) {
//           String batchId = BatchUtils.getBatchId(entry);
//           if (!BatchUtils.isSuccess(entry)) {
//             isSuccess = false;
//             BatchStatus status = BatchUtils.getBatchStatus(entry);
//             System.out.printf("%s failed (%s) %s", batchId, status.getReason(), status.getContent());
//
//           }
//
//            break;
//        }
        // 
//        logger.info(isSuccess ? "\nBatch operations successful." : "\nBatch operations failed");
        // System.out.printf("\n%s ms elapsed\n", System.currentTimeMillis() - startTime);


    }

    /**
    * A basic struct to store cell row/column information and the associated RnCn
    * identifier.
    * @author Josh Danziger
    */
    private static class CellAddress
    {
        public int row = 0;
        public int col = 0;
        public String idString ="";


        public CellAddress(){}

        /**
         * Constructs a CellAddress representing the specified {@code row} and
         * {@code col}.  The idString will be set in 'RnCn' notation.
         * @param row Row in spreadsheet
         * @param col column in spreadsheet
         */
        public CellAddress(int row, int col)
        {
          this.row = row;
          this.col = col;
          this.idString = String.format("R%sC%s", row, col);
        }

        public int getRow()
        {
            return row;
        }

        public void setRow(int row)
        {
            this.row = row;
        }

        public int getCol()
        {
            return col;
        }

        public void setCol(int col)
        {
            this.col = col;
        }

        public String getIdString()
        {
            return idString;
        }

        public void setIdString(String idString)
        {
            this.idString = idString;
        }
    }

    public void insertRow(String nameValuePairs, URL listFeedUrl) throws IOException, ServiceException
    {
        ListEntry newEntry = new ListEntry();

        // Split first by the commas between the different fields.
        for (String nameValuePair : nameValuePairs.split(",")) {

          // Then, split by the equal sign.
          String[] parts = nameValuePair.split("=", 2);
          String tag = parts[0]; // such as "name"
          String value = parts[1]; // such as "Fred"

          newEntry.getCustomElements().setValueLocal(tag, value);
        }

        ListEntry createEntry = spreadsheetService.insert(listFeedUrl, newEntry);
        logger.info("Created Entry: " + createEntry.getPlainTextContent());

    }

    public void insertRows(List<String> nameValuePairs, URL listFeedUrl) throws IOException, ServiceException
    {
        for(String nameValuePair : nameValuePairs)
        {
            insertRow(nameValuePair, listFeedUrl);
        }
    }

    public String getDocumentFeedUri(DocumentListEntry documentListEntry)
    {
        return ((MediaContent) documentListEntry.getContent()).getUri();
    }

    public URL getWorksheetURL(String spreadsheetName) throws IOException, ServiceException
    {
        URL worksheetURLtoFind = null;
        URL metafeedUrl = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
        SpreadsheetFeed feed = spreadsheetService.getFeed(metafeedUrl, SpreadsheetFeed.class);

        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        for (SpreadsheetEntry entry : spreadsheets)
        {
            logger.fine("\t" + entry.getTitle().getPlainText());
            if (spreadsheetName.equalsIgnoreCase(entry.getTitle().getPlainText()))
            {
                worksheetURLtoFind = entry.getWorksheetFeedUrl();
				break;
            }

        }
        return worksheetURLtoFind;
    }

    public WorksheetEntry getWorksheet(String spreadsheetName, String worksheetName) throws IOException, ServiceException
    {
        WorksheetEntry worksheetToFind = null;
        URL metafeedUrl = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");
        SpreadsheetFeed feed = spreadsheetService.getFeed(metafeedUrl, SpreadsheetFeed.class);

        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        for (SpreadsheetEntry entry : spreadsheets)
        {
            logger.info("\t" + entry.getTitle().getPlainText());
            if (spreadsheetName.equalsIgnoreCase(entry.getTitle().getPlainText()))
            {
                List<WorksheetEntry> worksheets = entry.getWorksheets();
                for (WorksheetEntry currentWorksheet : worksheets)
                {

                    logger.info("\t" + currentWorksheet.getTitle().getPlainText());

                    if (worksheetName.equalsIgnoreCase(currentWorksheet.getTitle().getPlainText()))
                    {
                        worksheetToFind = currentWorksheet;
                        break;
                    }
                }
            }

        }
        return worksheetToFind;
    }

  /**
   * Taken directly from Google Spreadsheet API because of Authorization Issue
   *
   * Gets all worksheet entries that are part of this spreadsheet.
   *
   * You must be online for this to work.
   *
   * @param spreadsheetEntry
     * @return the list of worksheet entries
     * @throws java.io.IOException
     * @throws com.google.gdata.util.ServiceException
   */
  public List<WorksheetEntry> getWorksheets(SpreadsheetEntry spreadsheetEntry)
      throws IOException, ServiceException
  {
    WorksheetFeed feed = spreadsheetService.getFeed(spreadsheetEntry.getWorksheetFeedUrl(),
        WorksheetFeed.class);
    return feed.getEntries();
  }

   public WorksheetEntry getWorksheet(List<WorksheetEntry> worksheets, String worksheetName)
   {
        logger.info("\n\n\n Get Worksheet URL***");

        WorksheetEntry worksheetToFind = new WorksheetEntry();

        for (WorksheetEntry currentWorksheet : worksheets)
        {

            logger.fine("\t" + currentWorksheet.getTitle().getPlainText());

            if (worksheetName.equalsIgnoreCase(currentWorksheet.getTitle().getPlainText()))
            {
                worksheetToFind = currentWorksheet;
                break;
            }
        }

        logger.info("\n\n\n Get Worksheet URL: " + worksheetToFind.getListFeedUrl());

        return worksheetToFind;
    }

    public com.google.gdata.data.docs.SpreadsheetEntry createSpreadsheet(String _title, DocumentListEntry _destinationFolderEntry, 
			String _defaultWorksheetName, int _rows, int _columns) throws IOException, ServiceException
    {
      com.google.gdata.data.docs.SpreadsheetEntry newEntry = new com.google.gdata.data.docs.SpreadsheetEntry();
      newEntry.setTitle(new PlainTextConstruct(_title));

      String destFolder = ((MediaContent) _destinationFolderEntry.getContent()).getUri();
	  URL destinationURL = new URL (destFolder);
	
	  com.google.gdata.data.docs.SpreadsheetEntry newSpreadsheetEntry = docsService.insert(destinationURL, newEntry);
	  //convert from Docs API Spreadsheet to Spreadsheet API Spreadsheet
	  WorksheetEntry worksheet = getSpreadsheetEntryFromDocumentListEntry(newSpreadsheetEntry).getDefaultWorksheet();
      worksheet.setTitle(new PlainTextConstruct(_defaultWorksheetName));
      worksheet.setRowCount(_rows);
      worksheet.setColCount(_columns);
	  worksheet.update();
	
      return newSpreadsheetEntry;
    }

	public com.google.gdata.data.docs.SpreadsheetEntry createSpreadsheet(String _title, URL _destinationURL, 
			String _defaultWorksheetName, int _rows, int _columns) throws IOException, ServiceException
    {
      com.google.gdata.data.docs.SpreadsheetEntry newEntry = new com.google.gdata.data.docs.SpreadsheetEntry();
      newEntry.setTitle(new PlainTextConstruct(_title));
	
	  com.google.gdata.data.docs.SpreadsheetEntry newSpreadsheetEntry = docsService.insert(_destinationURL, newEntry);
	  //convert from Docs API Spreadsheet to Spreadsheet API Spreadsheet
	  WorksheetEntry worksheet = getSpreadsheetEntryFromDocumentListEntry(newSpreadsheetEntry).getDefaultWorksheet();
      worksheet.setTitle(new PlainTextConstruct(_defaultWorksheetName));
      worksheet.setRowCount(_rows);
      worksheet.setColCount(_columns);
	  worksheet.update();
	
      return newSpreadsheetEntry;
    }


}