// Content script for ETA Invoice Exporter with exact data extraction matching the portal structure
class ETAContentScript {
  constructor() {
    this.invoiceData = [];
    this.allPagesData = [];
    this.totalCount = 0;
    this.currentPage = 1;
    this.totalPages = 1;
    this.isProcessingAllPages = false;
    this.progressCallback = null;
    this.init();
  }
  
  init() {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => this.scanForInvoices());
    } else {
      this.scanForInvoices();
    }
    
    this.setupMutationObserver();
  }
  
  setupMutationObserver() {
    this.observer = new MutationObserver((mutations) => {
      let shouldRescan = false;
      
      mutations.forEach((mutation) => {
        if (mutation.type === 'childList' && mutation.addedNodes.length > 0) {
          mutation.addedNodes.forEach((node) => {
            if (node.nodeType === Node.ELEMENT_NODE) {
              if (node.classList?.contains('ms-DetailsRow') || 
                  node.querySelector?.('.ms-DetailsRow') ||
                  node.classList?.contains('ms-List-cell') ||
                  node.classList?.contains('eta-pageNumber')) {
                shouldRescan = true;
              }
            }
          });
        }
      });
      
      if (shouldRescan && !this.isProcessingAllPages) {
        clearTimeout(this.rescanTimeout);
        this.rescanTimeout = setTimeout(() => this.scanForInvoices(), 1000);
      }
    });
    
    this.observer.observe(document.body, {
      childList: true,
      subtree: true
    });
  }
  
  scanForInvoices() {
    try {
      this.invoiceData = [];
      
      // Extract pagination info first
      this.extractPaginationInfo();
      
      // Find invoice rows using the exact selectors from the HTML
      const rows = document.querySelectorAll('.ms-DetailsRow[role="row"]');
      console.log(`ETA Exporter: Found ${rows.length} invoice rows on page ${this.currentPage}`);
      
      rows.forEach((row, index) => {
        const invoiceData = this.extractDataFromRow(row, index + 1);
        if (this.isValidInvoiceData(invoiceData)) {
          this.invoiceData.push(invoiceData);
        }
      });
      
      console.log(`ETA Exporter: Extracted ${this.invoiceData.length} valid invoices from page ${this.currentPage}`);
      
    } catch (error) {
      console.error('ETA Exporter: Error scanning for invoices:', error);
    }
  }
  
  extractPaginationInfo() {
    try {
      // Extract total count from pagination - exact selector from HTML
      const totalLabel = document.querySelector('.eta-pagination-totalrecordCount-label');
      if (totalLabel) {
        const match = totalLabel.textContent.match(/النتائج:\s*(\d+)/);
        if (match) {
          this.totalCount = parseInt(match[1]);
        }
      }
      
      // Extract current page - exact selector from HTML
      const currentPageBtn = document.querySelector('.eta-pageNumber.is-checked');
      if (currentPageBtn) {
        const pageLabel = currentPageBtn.querySelector('.ms-Button-label');
        if (pageLabel) {
          this.currentPage = parseInt(pageLabel.textContent) || 1;
        }
      }
      
      // Calculate total pages (10 items per page based on the HTML structure)
      const itemsPerPage = 10;
      this.totalPages = Math.ceil(this.totalCount / itemsPerPage);
      
      console.log(`ETA Exporter: Page ${this.currentPage} of ${this.totalPages}, Total: ${this.totalCount} invoices`);
      
    } catch (error) {
      console.warn('ETA Exporter: Error extracting pagination info:', error);
    }
  }
  
  extractDataFromRow(row, index) {
    // Extract data exactly as shown in the HTML structure
    const invoice = {
      index: index,
      pageNumber: this.currentPage,
      
      // Main invoice data from the table structure
      electronicNumber: '',
      internalNumber: '',
      issueDate: '',
      issueTime: '',
      documentType: '',
      documentVersion: '',
      totalAmount: '',
      currency: 'EGP',
      sellerName: '',
      sellerTaxNumber: '',
      buyerName: '',
      buyerTaxNumber: '',
      submissionId: '',
      status: '',
      
      // Additional calculated fields
      invoiceValue: '',
      vatAmount: '',
      totalInvoice: '',
      
      // Default values for missing fields
      taxDiscount: '0',
      sellerAddress: '',
      buyerAddress: '',
      purchaseOrderRef: '',
      purchaseOrderDesc: '',
      salesOrderRef: '',
      electronicSignature: 'موقع إلكترونياً',
      foodDrugGuide: '',
      externalLink: ''
    };
    
    try {
      // Get all cells in the row - exact structure from HTML
      const cells = row.querySelectorAll('.ms-DetailsRow-cell');
      
      if (cells.length === 0) {
        console.warn(`No cells found in row ${index}`);
        return invoice;
      }
      
      // Cell 0: Electronic Number and Internal Number (data-automation-key="uuid")
      const uuidCell = cells[0];
      if (uuidCell && uuidCell.getAttribute('data-automation-key') === 'uuid') {
        // Extract electronic number from the link
        const electronicLink = uuidCell.querySelector('.internalId-link a.griCellTitle');
        if (electronicLink) {
          invoice.electronicNumber = electronicLink.textContent?.trim() || '';
        }
        
        // Extract internal number from subtitle
        const internalNumberElement = uuidCell.querySelector('.griCellSubTitle');
        if (internalNumberElement) {
          invoice.internalNumber = internalNumberElement.textContent?.trim() || '';
        }
      }
      
      // Cell 1: Date and Time (data-automation-key="dateTimeReceived")
      const dateCell = cells[1];
      if (dateCell && dateCell.getAttribute('data-automation-key') === 'dateTimeReceived') {
        const dateElement = dateCell.querySelector('.griCellTitleGray');
        const timeElement = dateCell.querySelector('.griCellSubTitle');
        
        if (dateElement) {
          invoice.issueDate = dateElement.textContent?.trim() || '';
        }
        if (timeElement) {
          invoice.issueTime = timeElement.textContent?.trim() || '';
        }
      }
      
      // Cell 2: Document Type and Version (data-automation-key="typeName")
      const typeCell = cells[2];
      if (typeCell && typeCell.getAttribute('data-automation-key') === 'typeName') {
        const typeElement = typeCell.querySelector('.griCellTitleGray');
        const versionElement = typeCell.querySelector('.griCellSubTitle');
        
        if (typeElement) {
          invoice.documentType = typeElement.textContent?.trim() || '';
        }
        if (versionElement) {
          invoice.documentVersion = versionElement.textContent?.trim() || '';
        }
      }
      
      // Cell 3: Total Amount (data-automation-key="total")
      const totalCell = cells[3];
      if (totalCell && totalCell.getAttribute('data-automation-key') === 'total') {
        const totalElement = totalCell.querySelector('.griCellTitleGray');
        if (totalElement) {
          invoice.totalAmount = totalElement.textContent?.trim() || '';
          invoice.totalInvoice = invoice.totalAmount;
          invoice.invoiceValue = invoice.totalAmount;
        }
      }
      
      // Cell 4: Seller/Issuer Information (data-automation-key="issuerName")
      const issuerCell = cells[4];
      if (issuerCell && issuerCell.getAttribute('data-automation-key') === 'issuerName') {
        const sellerNameElement = issuerCell.querySelector('.griCellTitleGray');
        const sellerTaxElement = issuerCell.querySelector('.griCellSubTitle');
        
        if (sellerNameElement) {
          invoice.sellerName = sellerNameElement.textContent?.trim() || '';
        }
        if (sellerTaxElement) {
          invoice.sellerTaxNumber = sellerTaxElement.textContent?.trim() || '';
        }
      }
      
      // Cell 5: Buyer/Receiver Information (data-automation-key="receiverName")
      const receiverCell = cells[5];
      if (receiverCell && receiverCell.getAttribute('data-automation-key') === 'receiverName') {
        const buyerNameElement = receiverCell.querySelector('.griCellTitleGray');
        const buyerTaxElement = receiverCell.querySelector('.griCellSubTitle');
        
        if (buyerNameElement) {
          invoice.buyerName = buyerNameElement.textContent?.trim() || '';
        }
        if (buyerTaxElement) {
          invoice.buyerTaxNumber = buyerTaxElement.textContent?.trim() || '';
        }
      }
      
      // Cell 6: Submission ID (data-automation-key="submission")
      const submissionCell = cells[6];
      if (submissionCell && submissionCell.getAttribute('data-automation-key') === 'submission') {
        const submissionLink = submissionCell.querySelector('a.submissionId-link');
        if (submissionLink) {
          invoice.submissionId = submissionLink.textContent?.trim() || '';
          invoice.purchaseOrderRef = invoice.submissionId;
        }
      }
      
      // Cell 7: Status (data-automation-key="status")
      const statusCell = cells[7];
      if (statusCell && statusCell.getAttribute('data-automation-key') === 'status') {
        // Handle different status types exactly as in HTML
        const validRejectedDiv = statusCell.querySelector('.horizontal.valid-rejected');
        if (validRejectedDiv) {
          // Complex status: Valid → Rejected
          const validStatus = validRejectedDiv.querySelector('.status-Valid');
          const rejectedStatus = validRejectedDiv.querySelector('.status-Rejected');
          if (validStatus && rejectedStatus) {
            invoice.status = `${validStatus.textContent?.trim()} → ${rejectedStatus.textContent?.trim()}`;
          }
        } else {
          // Simple status
          const textStatus = statusCell.querySelector('.textStatus');
          if (textStatus) {
            invoice.status = textStatus.textContent?.trim() || '';
          } else {
            // Fallback to any text content
            invoice.status = statusCell.textContent?.trim() || '';
          }
        }
      }
      
      // Calculate VAT amount (assuming 14% VAT rate common in Egypt)
      if (invoice.totalAmount) {
        const totalValue = parseFloat(invoice.totalAmount.replace(/[,٬]/g, ''));
        if (!isNaN(totalValue)) {
          // Calculate VAT as 14% of the net amount
          invoice.vatAmount = (totalValue * 0.14 / 1.14).toFixed(2);
          invoice.invoiceValue = (totalValue - parseFloat(invoice.vatAmount)).toFixed(2);
        }
      }
      
      // Set default addresses if names are available
      if (invoice.sellerName && !invoice.sellerAddress) {
        invoice.sellerAddress = 'غير محدد';
      }
      if (invoice.buyerName && !invoice.buyerAddress) {
        invoice.buyerAddress = 'غير محدد';
      }
      
      // Generate external link for sharing
      if (invoice.electronicNumber) {
        invoice.externalLink = this.generateExternalLink(invoice);
      }
      
    } catch (error) {
      console.warn(`ETA Exporter: Error extracting data from row ${index}:`, error);
    }
    
    return invoice;
  }
  
  generateExternalLink(invoice) {
    if (!invoice.electronicNumber) {
      return '';
    }
    
    // Generate share link based on the submission ID or electronic number
    let shareId = '';
    if (invoice.submissionId && invoice.submissionId.length > 10) {
      shareId = invoice.submissionId;
    } else {
      // Generate a placeholder shareId based on electronic number
      shareId = invoice.electronicNumber.replace(/[^A-Z0-9]/g, '').substring(0, 26);
    }
    
    return `https://invoicing.eta.gov.eg/documents/${invoice.electronicNumber}/share/${shareId}`;
  }
  
  isValidInvoiceData(invoice) {
    return !!(invoice.electronicNumber || invoice.internalNumber || invoice.totalAmount);
  }
  
  async getAllPagesData(options = {}) {
    try {
      this.isProcessingAllPages = true;
      this.allPagesData = [];
      
      // First, scan current page to get pagination info
      this.scanForInvoices();
      
      if (this.totalPages <= 1) {
        // Only one page, return current data
        return {
          success: true,
          data: this.invoiceData,
          totalProcessed: this.invoiceData.length
        };
      }
      
      // Start from page 1
      await this.navigateToPage(1);
      
      // Process all pages
      for (let page = 1; page <= this.totalPages; page++) {
        try {
          // Update progress
          if (this.progressCallback) {
            this.progressCallback({
              currentPage: page,
              totalPages: this.totalPages,
              message: `جاري معالجة الصفحة ${page} من ${this.totalPages}...`
            });
          }
          
          // Wait for page to load and scan
          await this.waitForPageLoad();
          this.scanForInvoices();
          
          // Add current page data to all pages data
          this.allPagesData.push(...this.invoiceData);
          
          console.log(`ETA Exporter: Processed page ${page}, collected ${this.invoiceData.length} invoices`);
          
          // Navigate to next page if not the last page
          if (page < this.totalPages) {
            await this.navigateToNextPage();
            await this.delay(1500); // Wait between page transitions
          }
          
        } catch (error) {
          console.error(`Error processing page ${page}:`, error);
          // Continue with next page even if current page fails
        }
      }
      
      return {
        success: true,
        data: this.allPagesData,
        totalProcessed: this.allPagesData.length
      };
      
    } catch (error) {
      console.error('Error getting all pages data:', error);
      return { 
        success: false, 
        data: this.allPagesData,
        error: error.message 
      };
    } finally {
      this.isProcessingAllPages = false;
    }
  }
  
  async navigateToPage(pageNumber) {
    try {
      // Look for page number button using exact selector from HTML
      const pageButtons = document.querySelectorAll('.eta-pageNumber');
      
      // Find the specific page button
      for (const btn of pageButtons) {
        const label = btn.querySelector('.ms-Button-label');
        if (label && parseInt(label.textContent) === pageNumber) {
          btn.click();
          await this.waitForPageLoad();
          return true;
        }
      }
      
      return false;
    } catch (error) {
      console.error(`Error navigating to page ${pageNumber}:`, error);
      return false;
    }
  }
  
  async navigateToNextPage() {
    try {
      // Look for next page button using exact selector from HTML
      const nextButton = document.querySelector('[data-icon-name="ChevronRight"]')?.closest('button');
      
      if (nextButton && !nextButton.disabled && !nextButton.classList.contains('is-disabled')) {
        nextButton.click();
        return true;
      }
      
      return false;
    } catch (error) {
      console.error('Error navigating to next page:', error);
      return false;
    }
  }
  
  async waitForPageLoad() {
    // Wait for loading indicators to disappear
    await this.waitForCondition(() => {
      const loadingIndicators = document.querySelectorAll('.LoadingIndicator, .ms-Spinner');
      return loadingIndicators.length === 0 || 
             Array.from(loadingIndicators).every(el => el.style.display === 'none' || !el.offsetParent);
    }, 10000);
    
    // Wait for invoice rows to appear
    await this.waitForCondition(() => {
      const rows = document.querySelectorAll('.ms-DetailsRow[role="row"]');
      return rows.length > 0;
    }, 10000);
    
    // Additional wait to ensure data is fully loaded
    await this.delay(1000);
  }
  
  async waitForCondition(condition, timeout = 5000) {
    const startTime = Date.now();
    
    while (Date.now() - startTime < timeout) {
      if (condition()) {
        return true;
      }
      await this.delay(100);
    }
    
    return false;
  }
  
  delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  setProgressCallback(callback) {
    this.progressCallback = callback;
  }
  
  async getInvoiceDetails(invoiceId) {
    try {
      // Extract invoice details from the details view
      // This would be called when user clicks on an invoice link
      const details = await this.extractInvoiceDetailsFromPage(invoiceId);
      
      return {
        success: true,
        data: details
      };
    } catch (error) {
      console.error('Error getting invoice details:', error);
      return { 
        success: false, 
        data: [],
        error: error.message 
      };
    }
  }
  
  async extractInvoiceDetailsFromPage(invoiceId) {
    // This method extracts line items from the invoice details page
    // Based on the second image showing the detailed view with columns:
    // كود الصنف, إسم الكود, الوصف, الكمية, كود الوحدة, إسم الوحدة, السعر, القيمة, ضريبة القيمة المضافة, الإجمالي
    
    const details = [];
    
    try {
      // Look for the details table in the invoice view
      const detailsTable = document.querySelector('.ms-DetailsList, [data-automationid="DetailsList"]');
      
      if (detailsTable) {
        const rows = detailsTable.querySelectorAll('.ms-DetailsRow[role="row"]');
        
        rows.forEach((row, index) => {
          const cells = row.querySelectorAll('.ms-DetailsRow-cell');
          
          if (cells.length >= 10) { // Based on the 10 columns shown in the image
            const item = {
              itemCode: this.extractCellText(cells[0]) || '', // كود الصنف
              codeName: this.extractCellText(cells[1]) || '', // إسم الكود
              description: this.extractCellText(cells[2]) || '', // الوصف
              quantity: this.extractCellText(cells[3]) || '1', // الكمية
              unitCode: this.extractCellText(cells[4]) || 'EA', // كود الوحدة
              unitName: this.extractCellText(cells[5]) || 'قطعة', // إسم الوحدة
              unitPrice: this.extractCellText(cells[6]) || '0', // السعر
              totalValue: this.extractCellText(cells[7]) || '0', // القيمة
              vatAmount: this.extractCellText(cells[8]) || '0', // ضريبة القيمة المضافة
              totalWithVat: this.extractCellText(cells[9]) || '0' // الإجمالي
            };
            
            // Only add valid items (skip header rows)
            if (item.description && item.description !== 'الوصف') {
              details.push(item);
            }
          }
        });
      }
      
      // If no detailed items found, create a summary item
      if (details.length === 0) {
        const invoice = this.invoiceData.find(inv => inv.electronicNumber === invoiceId);
        if (invoice) {
          details.push({
            itemCode: 'SUMMARY',
            codeName: 'إجمالي',
            description: 'إجمالي الفاتورة',
            quantity: '1',
            unitCode: 'EA',
            unitName: 'قطعة',
            unitPrice: invoice.totalAmount || '0',
            totalValue: invoice.invoiceValue || invoice.totalAmount || '0',
            vatAmount: invoice.vatAmount || '0',
            totalWithVat: invoice.totalAmount || '0'
          });
        }
      }
      
    } catch (error) {
      console.error('Error extracting invoice details:', error);
    }
    
    return details;
  }
  
  extractCellText(cell) {
    if (!cell) return '';
    
    // Try different selectors for cell content
    const textElement = cell.querySelector('.griCellTitle, .griCellTitleGray, .ms-DetailsRow-cellContent') || cell;
    return textElement.textContent?.trim() || '';
  }
  
  getInvoiceData() {
    return {
      invoices: this.invoiceData,
      totalCount: this.totalCount,
      currentPage: this.currentPage,
      totalPages: this.totalPages
    };
  }
  
  cleanup() {
    if (this.observer) {
      this.observer.disconnect();
    }
    if (this.rescanTimeout) {
      clearTimeout(this.rescanTimeout);
    }
  }
}

// Initialize content script
const etaContentScript = new ETAContentScript();

// Listen for messages from popup
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  switch (request.action) {
    case 'getInvoiceData':
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    case 'getInvoiceDetails':
      etaContentScript.getInvoiceDetails(request.invoiceId)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true; // Indicates async response
      
    case 'getAllPagesData':
      // Set up progress callback
      if (request.options && request.options.progressCallback) {
        etaContentScript.setProgressCallback((progress) => {
          chrome.runtime.sendMessage({
            action: 'progressUpdate',
            progress: progress
          });
        });
      }
      
      etaContentScript.getAllPagesData(request.options)
        .then(result => sendResponse(result))
        .catch(error => sendResponse({ success: false, error: error.message }));
      return true; // Indicates async response
      
    case 'rescanPage':
      etaContentScript.scanForInvoices();
      sendResponse({
        success: true,
        data: etaContentScript.getInvoiceData()
      });
      break;
      
    default:
      sendResponse({ success: false, error: 'Unknown action' });
  }
  
  return true;
});

// Cleanup on page unload
window.addEventListener('beforeunload', () => {
  etaContentScript.cleanup();
});