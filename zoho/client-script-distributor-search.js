// Client Script: Distributor Search (Widget POPUP with Deluge Integration)
// Module: Sales Quotes
// Trigger: Custom Button "Distributor Search"
// Version: 1.8
// Updated: September 28, 2026
//   - Every value sent to the Deluge functions is now URL-encoded. The values go
//     out as query-string parameters, so an unencoded "&" (e.g. Almo category
//     "Mounts & Racks") split the request and Zoho rejected it with HTTP 400,
//     surfacing as an empty ZDKError.
//   - Products the widget already saved (Zoho_Product_Id set) skip the Deluge
//     call here; the widget reports those failures before it closes.
//   - Each Deluge call is wrapped so one failure is reported by SKU instead of
//     aborting the whole run.
//   - Vendor Direct Vendor_Lookup map also keyed by the Zoho manufacturer names
//     the widget actually sends (e.g. "Extron Electronics", not "Extron").
console.log("CLIENT SCRIPT VERSION 1.8 LOADED - ENCODED PARAMS");

// Vendor Direct manufacturers don't share one distributor Vendor record — each
// manufacturer is bought direct, so Vendor_Lookup must resolve to that
// manufacturer's own Vendors-module record. Matched case-insensitively against
// the manufacturer name the widget sends (the Zoho Manufacturers name); the
// short names are kept for older widget builds.
var VENDOR_DIRECT_VENDOR_MAP = {
  "Extron Electronics": { id: "5439147000009501601", name: "EXTRON ELECTRONICS" },
  "Extron": { id: "5439147000009501601", name: "EXTRON ELECTRONICS" },
  "Shure, Inc.": { id: "5439147000009504375", name: "Shure, Inc." },
  "Shure": { id: "5439147000009504375", name: "Shure, Inc." },
  "Crestron Electronic, Inc.": { id: "5439147000009501603", name: "Crestron Electronic, Inc." },
  "Crestron": { id: "5439147000009501603", name: "Crestron Electronic, Inc." },
  "Allen & Heath Limited": { id: "5439147000078957099", name: "Allen & Heath" },
  "Allen & Heath": { id: "5439147000078957099", name: "Allen & Heath" },
  "Listen Technologies": { id: "5439147000009504426", name: "Listen Technologies" },
  "LEA Professional": { id: "5439147000009504420", name: "LEA, LLC" },
};

function findVendorDirectVendor(manufacturerName) {
  if (!manufacturerName) return null;
  var target = String(manufacturerName).toLowerCase();
  for (var key in VENDOR_DIRECT_VENDOR_MAP) {
    if (key.toLowerCase() === target) return VENDOR_DIRECT_VENDOR_MAP[key];
  }
  return null;
}

// Deluge function per distributor. Almo and Teledynamics are listed under both
// their API name and their code name; the first one that responds is used.
var DELUGE_FUNCTIONS = {
  "Ingram Micro": ["createingramproduct"],
  "TD SYNNEX": ["createtdsynnexproduct"],
  "ADI Global": ["createadiglobalproduct"],
  "Vendor Direct": ["createvendordirectproduct"],
  "Almo": ["createalmoproduct", "create_almo_product"],
  "Teledynamics": ["createteledynamicsproduct", "create_teledynamics_product"],
};

// Values are sent as query-string parameters, so each one must be encoded
// exactly once. The Deluge functions urlDecode the free-text fields again,
// which is harmless for already-decoded text.
function enc(value) {
  return encodeURIComponent(value === null || value === undefined ? "" : String(value));
}

function describeZdkError(err) {
  if (!err) return "unknown error";
  if (err.message) return err.message;
  if (err.code) return err.code;
  var s = err.toString();
  return s === "ZDKError" ? "request rejected by Zoho (ZDKError)" : s;
}

// Calls the first function name that succeeds. Returns the ZDK response, or
// throws the last error if every name failed.
function executeDeluge(functionNames, params) {
  var lastErr = null;
  for (var n = 0; n < functionNames.length; n++) {
    try {
      console.log("CS: Calling Deluge function: " + functionNames[n]);
      return ZDK.Apps.CRM.Functions.execute(functionNames[n], params);
    } catch (err) {
      console.error("CS: Deluge call failed for " + functionNames[n] + ": " + describeZdkError(err));
      lastErr = err;
    }
  }
  throw lastErr;
}

// =============================================================================
// STEP 1: Pre-fetch Manufacturers from Zoho CRM
// =============================================================================
// If fetch fails, widget still functions with empty manufacturer list.
// NOTE: No setTimeout/Promise.race — Zoho's sandboxed runtime does not support setTimeout.

var fetchManufacturers = new Promise(function (resolve) {
  try {
    // ZDK returns max perPage records per call — paginate to fetch all
    var allRecords = [];
    var page = 1;
    var perPage = 200;
    var pageRecords;
    do {
      pageRecords = ZDK.Apps.CRM.Manufacturers.fetch(page, perPage, "Name", "asc");
      if (pageRecords && pageRecords.length > 0) {
        allRecords = allRecords.concat(pageRecords);
      }
      page++;
    } while (pageRecords && pageRecords.length === perPage);
    console.log("DEBUG: Fetched " + allRecords.length + " manufacturers from Zoho CRM (paginated, " + (page - 1) + " pages)");
    resolve(
      allRecords.map(function (m) {
        return { id: m.id, name: m.Name };
      }),
    );
  } catch (err) {
    console.error("CS: ERROR at manufacturer fetch: " + describeZdkError(err));
    resolve([]);
  }
});

fetchManufacturers
  .then(function (manufacturersList) {
    manufacturersList = manufacturersList || [];
    console.log("DEBUG: Manufacturers ready for widget:", manufacturersList.length);

    // =============================================================================
    // STEP 2: Open popup with widget
    // =============================================================================

    console.log("CS: Opening widget popup");
    var response = ZDK.Client.openPopup(
      {
        api_name: "Distributor_Search",
        type: "widget",
        header: "",
        animation_type: 1,
        close_icon: false,
        close_on_escape: true,
        height: "1000px",
        width: "1200px",
        left: "center",
      },
      {
        data: {
          action: "search_products",
          manufacturers: manufacturersList,
        },
        wait: true,
      },
    );

    // =============================================================================
    // STEP 3: Validate response from widget
    // =============================================================================

    console.log("CS: Widget response received");
    console.log("CS: Widget response products:", response && response.products ? JSON.stringify(response.products) : "none");

    if (!response || response.cancelled || !response.products || response.products.length === 0) {
      console.log("DEBUG: Exiting early - no valid products in response");
      return;
    }

    var products = response.products;

    // =============================================================================
    // STEP 4: Create/update each product (skipped when the widget already did it)
    // =============================================================================

    var processedProducts = [];
    var errors = [];

    // Show loader to prevent client script timeout during Deluge calls (per Kaizen #139)
    // NOTE: Cannot use ZDK.Client.showMessage() while showLoader() is active (causes ZDKError)
    ZDK.Client.showLoader({ type: "page", template: "spinner", message: "Processing products..." });

    for (var i = 0; i < products.length; i++) {
      var product = products[i];
      console.log("CS: Processing product " + (i + 1) + " of " + products.length + " (" + product.Product_Code + ")");

      // Widget already created/updated this product before closing
      if (product.Zoho_Product_Id) {
        console.log("CS: Using product saved by widget - product_id: " + product.Zoho_Product_Id);
        processedProducts.push({
          product_id: product.Zoho_Product_Id,
          product_name: product.Product_Name,
          manufacturer_id: product.Zoho_Manufacturer_Id || null,
          manufacturer_name: product.Manufacturer,
          msrp: product.MSRP || 0,
          customer_price: product.Customer_Price || 0,
          action: product.Zoho_Action || "updated",
          customer_discount: product.Customer_Discount || 0,
          _sourceIndex: i,
        });
        continue;
      }

      // Every value is encoded exactly once (see enc()).
      var productData = {
        manufacturer_part_number: enc(product.Product_Code),
        manufacturer_name: enc(product.Manufacturer),
        product_name: enc(product.Product_Name),
        msrp: product.MSRP ? enc(product.MSRP) : "",
        customer_price: product.Customer_Price ? enc(product.Customer_Price) : "",
        description: enc(product.Description),
        upc: enc(product.UPC),
        last_sync_source: enc(product.Last_Sync_Source || "Ingram Micro"),
        // Ingram Micro fields
        ingram_micro_sku: enc(product.Ingram_Micro_SKU),
        category: enc(product.Category),
        subcategory: enc(product.Subcategory),
        im_product_type: enc(product.IM_Product_Type),
        // TD Synnex fields
        tdsynnex_sku: enc(product.TDSynnex_SKU),
        category_level_1: enc(product.Category_Level_1),
        category_level_2: enc(product.Category_Level_2),
        category_level_3: enc(product.Category_Level_3),
        // ADI Global fields
        adi_sku: enc(product.ADI_SKU),
        category_1: enc(product.ADI_Category_1),
        category_2: enc(product.ADI_Category_2),
        // Almo fields
        almo_sku: enc(product.Almo_SKU),
        almo_category_1: enc(product.Almo_Category_1),
        almo_category_2: enc(product.Almo_Category_2),
        // Teledynamics fields
        teledynamics_pn: enc(product.Teledynamics_SKU),
        teledynamics_category_1: enc(product.Teledynamics_Category_1),
        teledynamics_category_2: enc(product.Teledynamics_Category_2),
        // Vendor Direct fields
        vendor_direct_category: enc(product.Vendor_Direct_Category),
        // Shared fields
        unspsc_commodity: enc(product.UNSPSC_Commodity),
        kit_or_standalone: enc(product.Kit_or_Standalone),
        replacement_sku: enc((product.Replacement_SKU || "").replace(/[\r\n]/g, "")),
      };

      var source = product.Last_Sync_Source || "Ingram Micro";
      var functionNames = DELUGE_FUNCTIONS[source] || DELUGE_FUNCTIONS["Ingram Micro"];
      console.log("CS: Deluge productData: " + JSON.stringify(productData));

      var functionResponse;
      try {
        functionResponse = executeDeluge(functionNames, productData);
      } catch (err) {
        errors.push(product.Product_Code + ": " + describeZdkError(err));
        continue;
      }
      console.log("CS: Deluge response: " + JSON.stringify(functionResponse));

      if (!functionResponse) {
        errors.push(product.Product_Code + ": No response from function");
        continue;
      }
      if (functionResponse.code === "error") {
        console.error("CS: ERROR at Deluge function execution: " + functionResponse.message);
        errors.push(product.Product_Code + ": " + functionResponse.message);
        continue;
      }
      if (!functionResponse.details || !functionResponse.details.output) {
        console.error("CS: ERROR at response validation: missing output for " + product.Product_Code);
        errors.push(product.Product_Code + ": No output from function");
        continue;
      }

      var result;
      try {
        result = JSON.parse(functionResponse.details.output);
        console.log("CS: Parsed Deluge result: " + JSON.stringify(result));
      } catch (e) {
        console.error("CS: ERROR at JSON parse: " + e.toString() + " | raw output: " + functionResponse.details.output);
        errors.push(product.Product_Code + ": Parse error - " + e.toString());
        continue;
      }

      if (result.success === true && result.product_id) {
        processedProducts.push({
          product_id: result.product_id,
          product_name: product.Product_Name || result.product_name,
          manufacturer_id: result.manufacturer_id,
          manufacturer_name: product.Manufacturer,
          msrp: product.MSRP || 0,
          customer_price: product.Customer_Price || 0,
          action: result.action_taken,
          customer_discount: product.Customer_Discount || 0,
          _sourceIndex: i,
        });
      } else {
        var errorMsg = result.error || "Failed";
        if (result.details) {
          errorMsg += " - " + result.details;
        }
        console.error("CS: ERROR at Deluge result check: " + product.Product_Code + " - " + errorMsg);
        errors.push(product.Product_Code + ": " + errorMsg);
      }
    }

    ZDK.Client.hideLoader();

    // =============================================================================
    // STEP 5: Check if we have products to add
    // =============================================================================

    console.log("CS: Processed " + processedProducts.length + ", errors " + errors.length + ": " + JSON.stringify(errors));

    if (processedProducts.length === 0) {
      if (errors.length > 0) {
        ZDK.Client.showMessage("Failed to process products: " + errors.join(", "), "error");
      }
      return;
    }

    // =============================================================================
    // STEP 6: Add products to Quoted_Items subform
    // =============================================================================

    var quotedItemsSubform = ZDK.Page.getSubform("Quoted_Items");

    // Find the first empty row
    var insertIndex = 0;
    for (var j = 0; j < 100; j++) {
      try {
        var row = quotedItemsSubform.getRow(j);
        var rowValues = row.getValues();
        if (!rowValues.Product_Name || !rowValues.Product_Name.id) {
          insertIndex = j;
          break;
        }
      } catch (e) {
        insertIndex = j;
        break;
      }
    }

    var newRows = [];
    for (var k = 0; k < processedProducts.length; k++) {
      var proc = processedProducts[k];
      // _sourceIndex references the original product (handles gaps from failed calls)
      var srcProduct = products[proc._sourceIndex];

      var formattedProductName = proc.product_name + " (" + srcProduct.Product_Code + ")";

      // Line_Item_Cost MUST be per-unit cost: Zoho multiplies Line_Item_Cost × Quantity internally.
      var unitCost = proc.customer_price || 0;
      var qty = srcProduct.Quantity || 1;
      var discountPct = srcProduct.Customer_Discount || 0;
      var discountDollar = proc.msrp && discountPct ? parseFloat((proc.msrp * discountPct / 100).toFixed(2)) : 0;

      var newRow = {
        Product_Name: {
          id: proc.product_id,
          name: formattedProductName,
        },
        List_Price: proc.msrp,
        Unit_Cost_1: unitCost,
        Quantity: qty,
        Discount: discountDollar,
      };

      if (proc.manufacturer_id) {
        newRow.Manufacturer_Lookup = {
          id: proc.manufacturer_id,
          name: proc.manufacturer_name,
        };
      }

      // Vendor_Lookup based on distributor source
      if (srcProduct.Last_Sync_Source === "TD SYNNEX") {
        newRow.Vendor_Lookup = { name: "TD Synnex" };
      } else if (srcProduct.Last_Sync_Source === "ADI Global") {
        newRow.Vendor_Lookup = { name: "ADI Global" };
      } else if (srcProduct.Last_Sync_Source === "Almo") {
        // No Vendor record is named exactly "Almo" — the CRM record is "Exertis Almo"
        newRow.Vendor_Lookup = { id: "5439147000009501602", name: "Exertis Almo" };
      } else if (srcProduct.Last_Sync_Source === "Teledynamics") {
        newRow.Vendor_Lookup = { name: "Teledynamics" };
      } else if (srcProduct.Last_Sync_Source === "Vendor Direct") {
        var vdVendor = findVendorDirectVendor(proc.manufacturer_name);
        if (vdVendor) {
          newRow.Vendor_Lookup = { id: vdVendor.id, name: vdVendor.name };
        } else {
          console.warn("CS: No Vendor_Lookup mapping for Vendor Direct manufacturer: " + proc.manufacturer_name);
        }
      } else {
        newRow.Vendor_Lookup = { name: "Ingram Micro" };
      }

      console.log("CS: Subform row " + k + ": " + JSON.stringify(newRow));
      newRows.push(newRow);
    }

    console.log("CS: Inserting rows to Quoted_Items, count: " + newRows.length + ", at index: " + insertIndex);
    try {
      var insertResult = quotedItemsSubform.insertRows(newRows, insertIndex);
      console.log("CS: insertRows result: " + JSON.stringify(insertResult));
    } catch (err) {
      console.error("CS: ERROR at insertRows: " + describeZdkError(err) + " | rows: " + JSON.stringify(newRows));
      ZDK.Client.showMessage(
        "Products were saved but could not be added to the quote (" + describeZdkError(err) +
          "). Reopen Distributor Search and use Restore to try again.",
        "error",
      );
      return;
    }

    // =============================================================================
    // STEP 7: Show result message
    // =============================================================================

    var created = 0;
    var updated = 0;
    var existed = 0;
    for (var m = 0; m < processedProducts.length; m++) {
      if (processedProducts[m].action === "created") {
        created++;
      } else if (processedProducts[m].action === "updated") {
        updated++;
      } else if (processedProducts[m].action === "exists") {
        existed++;
      }
    }

    var message = processedProducts.length + " product(s) added to quote";
    if (created > 0 || updated > 0 || existed > 0) {
      message += " (" + created + " new, " + updated + " updated";
      if (existed > 0) {
        message += ", " + existed + " already existed - no fields updated";
      }
      message += ")";
    }
    if (errors.length > 0) {
      message += ". " + errors.length + " failed: " + errors.join(", ");
    }

    ZDK.Client.showMessage(message, errors.length > 0 ? "warning" : "success");
  })
  .catch(function (err) {
    ZDK.Client.hideLoader();
    // widget_closed is not an error - user clicked X to close
    if (err && err.toString().indexOf("widget_closed") !== -1) {
      console.log("CS: Widget closed by user (not an error)");
      return;
    }
    console.error("CS: ERROR at main promise chain: " + describeZdkError(err));
    try { console.error("CS: Error full JSON: " + JSON.stringify(err)); } catch (e2) { console.error("CS: Error not serializable"); }
    ZDK.Client.showMessage("Error: " + describeZdkError(err), "error");
  });
