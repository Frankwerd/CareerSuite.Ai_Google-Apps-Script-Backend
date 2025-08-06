// DocumentService.gs
// This file is responsible for creating and formatting the final resume Google Document.
// It uses a template document, populates placeholders with data from the resumeObject,
// and applies specific styling to different resume sections and elements.
// It relies on global constants defined in 'Constants.gs' (e.g., RESUME_TEMPLATE_DOC_ID).


// --- UTILITY FUNCTIONS ---

/**
 * Escapes special characters in a string for use in a regular expression.
 * If not already defined globally, this will define it.
 * @param {string} s The string to escape.
 * @return {string} The escaped string.
 */
if (typeof RegExp.escape !== 'function') {
  RegExp.escape = function(s) {
    if (typeof s !== 'string') return ''; // Ensure it's a string
    return s.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
  };
}

/**
 * Finds and replaces placeholder text in the document body with minimal styling impact,
 * primarily relying on the body.replaceText method. Includes detailed logging.
 * @param {GoogleAppsScript.Document.Body} body The document body element.
 * @param {string} placeholder The exact text placeholder to find (e.g., "{SUMMARY_CONTENT}").
 * @param {string} textToInsert The text to replace the placeholder with. Null or undefined becomes an empty string.
 */
function findAndReplaceText_minimal(body, placeholder, textToInsert) {
  const searchPattern = RegExp.escape(placeholder);
  const effectiveText = (textToInsert === null || textToInsert === undefined) ? "" : String(textToInsert);

  Logger.log(`>>> findAndReplaceText_minimal: Attempting to replace "${placeholder}" (escaped: "${searchPattern}") with text of length ${effectiveText.length}.`);
  if (effectiveText.length > 0) {
    // Logger.log_DEBUG(`    Preview of text for "${placeholder}": "${effectiveText.substring(0, 100)}..."`);
  } else {
    Logger.log(`    Text for "${placeholder}" is empty.`);
  }

  const foundRangeBefore = body.findText(searchPattern);
  if (foundRangeBefore) {
    Logger.log(`    SUCCESS (Before Replace): Found "${placeholder}" via findText. Element type: ${foundRangeBefore.getElement().getType()}, Parent type: ${foundRangeBefore.getElement().getParent().getType()}`);
  } else {
    Logger.log(`    FAILURE (Before Replace): Did NOT find "${placeholder}" via findText for body.replaceText.`);
  }

  body.replaceText(searchPattern, effectiveText);

  const foundRangeAfter = body.findText(searchPattern);
  if (foundRangeAfter) {
    Logger.log(`    WARN (After Replace): Placeholder "${placeholder}" STILL FOUND after replaceText attempt. Check template or if text was empty.`);
  } else {
    Logger.log(`    SUCCESS (After Replace): Placeholder "${placeholder}" NOT found after replaceText (assumed replaced or not found initially).`);
  }
  Logger.log(`>>> findAndReplaceText_minimal: Finished for "${placeholder}".`);
}


// --- ATTRIBUTE HANDLING HELPERS ---

/**
 * Sanitizes a raw attributes object (often from getAttributes()) to ensure values,
 * especially enums like HORIZONTAL_ALIGNMENT, are in the correct format for setAttributes().
 * V7 includes detailed logging for HorizontalAlignment processing.
 * @param {Object} rawAttrs The raw attributes object.
 * @return {Object} A new object with sanitized attributes.
 */
function sanitizeBaseAttributes(rawAttrs) {
  // Logger.log_DEBUG("--- SANITIZE_BASE_ATTRIBUTES_V7 ---"); // Less verbose for normal runs
  if (!rawAttrs) { return {}; }
  const sanitizedAttrs = {};
  // Define allowed DocumentApp.Attribute enums that this function will process
  const allowedAttributeEnums = [
    DocumentApp.Attribute.FONT_FAMILY, DocumentApp.Attribute.FONT_SIZE,
    DocumentApp.Attribute.FOREGROUND_COLOR, DocumentApp.Attribute.BACKGROUND_COLOR,
    DocumentApp.Attribute.BOLD, DocumentApp.Attribute.ITALIC, DocumentApp.Attribute.UNDERLINE,
    DocumentApp.Attribute.STRIKETHROUGH, DocumentApp.Attribute.LINE_SPACING,
    DocumentApp.Attribute.SPACING_BEFORE, DocumentApp.Attribute.SPACING_AFTER,
    DocumentApp.Attribute.INDENT_START, DocumentApp.Attribute.INDENT_END,
    DocumentApp.Attribute.HORIZONTAL_ALIGNMENT
  ];

  for (const keyStringFromRawAttrs in rawAttrs) {
    let matchingEnumKey = null;
    // Find the corresponding DocumentApp.Attribute enum for the string key
    for (const attrEnum of allowedAttributeEnums) {
      if (attrEnum.toString() === keyStringFromRawAttrs) {
        matchingEnumKey = attrEnum;
        break;
      }
    }

    if (matchingEnumKey) {
      let value = rawAttrs[keyStringFromRawAttrs];
      // Special handling for HORIZONTAL_ALIGNMENT to convert string values to enums
      if (matchingEnumKey === DocumentApp.Attribute.HORIZONTAL_ALIGNMENT) {
        // Logger.log_DEBUG(`  SanitizeV7: Processing HORIZONTAL_ALIGNMENT. Raw value type: ${typeof value}, value: ${value}`);
        if (typeof value === 'string') {
          const upperValue = value.toUpperCase();
          switch (upperValue) {
            case "LEFT": value = DocumentApp.HorizontalAlignment.LEFT; break;
            case "CENTER": value = DocumentApp.HorizontalAlignment.CENTER; break;
            case "RIGHT": value = DocumentApp.HorizontalAlignment.RIGHT; break;
            case "JUSTIFY": value = DocumentApp.HorizontalAlignment.JUSTIFY; break;
            default:
              Logger.log(`WARN (sanitizeBaseAttributes): Unknown HORIZONTAL_ALIGNMENT string: "${value}". Defaulting to LEFT.`);
              value = DocumentApp.HorizontalAlignment.LEFT;
          }
          // Logger.log_DEBUG(`    SanitizeV7: HORIZONTAL_ALIGNMENT converted to enum: ${value ? value.toString() : 'null'}`);
        } else if (value !== null && typeof value === 'object' && Object.values(DocumentApp.HorizontalAlignment).includes(value)) {
          // Logger.log_DEBUG(`    SanitizeV7: HORIZONTAL_ALIGNMENT is already a valid enum: ${value.toString()}.`);
        } else if (value !== null) {
          Logger.log(`WARN (sanitizeBaseAttributes): Invalid HORIZONTAL_ALIGNMENT value type: ${typeof value}. Value: ${value}. Defaulting to LEFT.`);
          value = DocumentApp.HorizontalAlignment.LEFT;
        } else {
          // Logger.log_DEBUG(`    SanitizeV7: HORIZONTAL_ALIGNMENT raw value is null.`);
        }
      }
      // Add to sanitized attributes if value is not null (null attributes clear existing ones)
      if (value !== null) {
        sanitizedAttrs[matchingEnumKey] = value;
      }
    }
  }
  // Logger.log_DEBUG(`  SanitizeV7 Final Result: ${JSON.stringify(sanitizedAttrs)}`);
  return sanitizedAttrs;
}

/**
 * Creates a shallow copy of an attributes object. This is particularly useful for
 * preserving DocumentApp.Attribute enum keys, as direct JSON.parse(JSON.stringify(attrs))
 * would convert enum keys to strings.
 * @param {Object} originalAttrs The attributes object to copy.
 * @return {Object} A new attributes object with the same key-value pairs.
 */
function copyAttributesPreservingEnums(originalAttrs) {
  const newAttrs = {};
  if (!originalAttrs) return newAttrs;
  for (const key in originalAttrs) {
    // Direct assignment preserves enum keys if they are actual enum objects.
    newAttrs[key] = originalAttrs[key];
  }
  return newAttrs;
}


// --- BLOCK POPULATION HELPER ---

/**
 * Populates a block placeholder in the document body with a list of items.
 * It finds the placeholder, clears it, extracts its base styling,
 * and then calls a specific renderItemFunction for each item in itemsData.
 * @param {GoogleAppsScript.Document.Body} body The document body.
 * @param {string} placeholder The text placeholder to find (e.g., "{EXPERIENCE_JOBS}").
 * @param {Array<Object>} itemsData An array of data objects to render.
 * @param {function} renderItemFunction The function responsible for rendering a single item
 *        (e.g., renderExperienceJob, renderEducationItem).
 */
function populateBlockPlaceholder(body, placeholder, itemsData, renderItemFunction) {
  Logger.log(`>>> populateBlockPlaceholder CALLED for placeholder: "${placeholder}"`);
  // Logger.log_DEBUG(`    itemsData length: ${itemsData ? itemsData.length : 'null/undefined'}. Is renderItemFunction a function? ${typeof renderItemFunction === 'function'}`);
  // if (itemsData && itemsData.length > 0) { // DEBUG
  //   Logger.log_DEBUG(`    Preview of first item in itemsData for "${placeholder}": ${JSON.stringify(itemsData[0]).substring(0, 250)}...`);
  // }

  const searchPattern = RegExp.escape(placeholder);
  const placeHolderRange = body.findText(searchPattern);

  if (!placeHolderRange) {
    Logger.log(`    FAILURE (populateBlockPlaceholder): Placeholder "${placeholder}" NOT FOUND in document body.`);
    return; // Exit if placeholder not found
  }
  // Logger.log_DEBUG(`    SUCCESS: Found placeholderRange for "${placeholder}"`);

  let currentDocElement = placeHolderRange.getElement();
  // Ensure we are working with a Paragraph element for styling and clearing
  if (!currentDocElement || currentDocElement.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    // Logger.log_DEBUG(`    WARN: Placeholder "${placeholder}" element type is ${currentDocElement ? currentDocElement.getType() : 'null'}. Trying parent...`);
    if (currentDocElement && currentDocElement.getParent().getType() === DocumentApp.ElementType.PARAGRAPH) {
        currentDocElement = currentDocElement.getParent();
    } else {
        Logger.log(`    ERROR (populateBlockPlaceholder): Placeholder "${placeholder}" or its parent is not a Paragraph. Cannot proceed with styling/clearing correctly.`);
        return;
    }
  }
  
  const placeholderPara = currentDocElement.asParagraph();
  const baseAttrsFromPlaceholder = placeholderPara.getAttributes(); // Get attributes BEFORE clearing text

  placeholderPara.clear(); // Clear the placeholder text (e.g., "{EXPERIENCE_JOBS}")

  // Adjust spacing of the (now empty) paragraph that held the placeholder to avoid unwanted gaps.
  placeholderPara.setSpacingAfter(0);
  placeholderPara.setSpacingBefore(0);
  // Logger.log_DEBUG(`    populateBlockPlaceholder: Adjusted SPACING_AFTER/BEFORE on cleared placeholder paragraph for "${placeholder}"`);
  // Logger.log_DEBUG(`    Base attributes inherited from "${placeholder}": ${JSON.stringify(baseAttrsFromPlaceholder)}`);

  if (itemsData && itemsData.length > 0) {
    const cleanBaseParaAttrs = sanitizeBaseAttributes(baseAttrsFromPlaceholder);
    // Logger.log_DEBUG(`    Cleaned base attributes for items in "${placeholder}": ${JSON.stringify(cleanBaseParaAttrs)}`);
    Logger.log(`    Rendering ${itemsData.length} items for "${placeholder}"...`);

    for (let i = 0; i < itemsData.length; i++) {
      // Logger.log_DEBUG(`        LOOP ITERATION ${i} for "${placeholder}". Calling renderItemFunction for item (preview): ${JSON.stringify(itemsData[i]).substring(0,100)}...`);
      // `currentDocElement` is passed as `insertAfterEl` to the render functions.
      // For the first item (i=0), this is the (cleared) placeholderPara itself.
      currentDocElement = renderItemFunction(itemsData[i], body, currentDocElement, cleanBaseParaAttrs, i);
      if (!currentDocElement) {
        Logger.log(`        WARN (populateBlockPlaceholder): renderItemFunction for item ${i} of "${placeholder}" returned null. Breaking loop.`);
        break;
      }
    }
    Logger.log(`    Finished rendering items for "${placeholder}".`);
  } else {
    Logger.log(`    No itemsData (or empty array) for "${placeholder}". Nothing to render in this block.`);
  }
  Logger.log(`>>> populateBlockPlaceholder ENDED for placeholder: "${placeholder}"`);
}


// --- INDIVIDUAL SECTION/ITEM RENDER FUNCTIONS ---
// These functions are responsible for creating document elements for each specific resume section type.

/**
 * Renders a single education item into the document body.
 * @param {Object} edu The education item data object.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl The element after which new content should be inserted.
 * @param {Object} baseAttrsFromPlaceholder The base paragraph attributes inherited from the placeholder.
 * @param {number} index The index of this item in its list (0 for the first item).
 * @return {GoogleAppsScript.Document.Element} The last document element inserted for this item.
 */
function renderEducationItem(edu, bodyRef, insertAfterEl, baseAttrsFromPlaceholder, index) {
  Logger.log(`  renderEducationItem (Standard): Index ${index}, Institution: "${edu.institution || 'N/A'}"`);
  let currentLastInsertedElement = insertAfterEl;
  let insertionIndex;

  try {
    insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1;
  } catch (e) {
    Logger.log(`    renderEducationItem: Error getting child index. Defaulting to end. ${e.message}`);
    insertionIndex = bodyRef.getNumChildren();
  }

  const itemLineAttrs = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);

  // Spacing between education items
  if (index > 0) {
    const spacingPara = bodyRef.insertParagraph(insertionIndex, "");
    let spacingParaStyle = {}; 
    spacingParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = 4; 
    spacingParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    spacingParaStyle[DocumentApp.Attribute.LINE_SPACING] = 0.8; 
    spacingPara.setAttributes(spacingParaStyle);
    currentLastInsertedElement = spacingPara;
    insertionIndex++;
  }
  
  itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0) ? 0 : 1;
  itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1; 
  itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;

  // Line 1: Institution (Bold)
  const instText = (edu.institution || "Unnamed Institution").trim();
  if (instText) {
    const instPara = bodyRef.insertParagraph(insertionIndex, "");
    instPara.setAttributes(itemLineAttrs);
    instPara.appendText(instText).setBold(true);
    currentLastInsertedElement = instPara;
    insertionIndex++;
  }

  // Line 2: Dates (Italic, on a new line)
  let gradDateText = (edu.endDate || "").trim();
  if (!gradDateText || gradDateText.toLowerCase() === "present") {
    if (edu.expectedGraduation && edu.expectedGraduation.trim()) {
      gradDateText = edu.expectedGraduation.trim() + " (Expected)";
    } else if (edu.startDate && edu.startDate.trim()) {
      gradDateText = `${edu.startDate.trim()} – Present`;
    } else if (!gradDateText) {
        gradDateText = "Present";
    }
  } else if (edu.startDate && edu.startDate.trim()){
      gradDateText = `${edu.startDate.trim()} – ${gradDateText}`;
  }
  if (!gradDateText.trim() && instText) gradDateText = "Dates N/A"; // Only add if inst was there

  if (gradDateText && gradDateText !== "Dates N/A") {
    const datePara = bodyRef.insertParagraph(insertionIndex, gradDateText);
    let dateParaAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    dateParaAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to institution line
    dateParaAttrs[DocumentApp.Attribute.ITALIC] = true;
    datePara.setAttributes(dateParaAttrs);
    currentLastInsertedElement = datePara;
    insertionIndex++;
  }

  // Line 3: Degree [, Location (Italic)]
  const degreeText = (edu.degree || "").trim();
  const locationText = (edu.location || "").trim();
  if (degreeText || locationText) {
    const degreeLocPara = bodyRef.insertParagraph(insertionIndex, "");
    let degreeLineSpecificAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    degreeLineSpecificAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    degreeLocPara.setAttributes(degreeLineSpecificAttrs);

    if (degreeText) degreeLocPara.appendText(degreeText);
    if (locationText) {
      degreeLocPara.appendText(degreeText ? ", " : ""); 
      const locationStartIndex = degreeLocPara.getText().length;
      degreeLocPara.appendText(locationText);
      degreeLocPara.editAsText().setItalic(locationStartIndex, degreeLocPara.getText().length - 1, true);
    }
    currentLastInsertedElement = degreeLocPara;
    insertionIndex++;
  }
  
  // Line 4: GPA (Optional, smaller font)
  const gpaTextVal = (edu.gpa || "").trim();
  if (gpaTextVal) {
    const gpaPara = bodyRef.insertParagraph(insertionIndex, "");
    let gpaAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    gpaAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    gpaAttrs[DocumentApp.Attribute.FONT_SIZE] = Math.max(8, (itemLineAttrs[DocumentApp.Attribute.FONT_SIZE] || 10) - 1); 
    gpaPara.setAttributes(gpaAttrs);
    gpaPara.appendText("GPA: " + gpaTextVal);
    currentLastInsertedElement = gpaPara;
    insertionIndex++;
  }

  // Line 5: Relevant Coursework (Optional)
  const coursework = edu.relevantCoursework || [];
  if (Array.isArray(coursework) && coursework.length > 0) {
    const courseworkText = coursework.join(", ").trim();
    if (courseworkText) {
      const courseworkPara = bodyRef.insertParagraph(insertionIndex, "");
      let courseworkAttrs = copyAttributesPreservingEnums(itemLineAttrs);
      courseworkAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
      courseworkPara.setAttributes(courseworkAttrs);
      courseworkPara.appendText("Relevant Coursework: ").setItalic(true); 
      courseworkPara.appendText(courseworkText); 
      currentLastInsertedElement = courseworkPara;
    }
  }
  
  Logger.log(`  renderEducationItem (Standard): Finished for "${edu.institution || 'N/A'}"`);
  return currentLastInsertedElement;
}

/**
 * Renders a single experience/job item into the document body.
 * Includes company, title/location, dates, and responsibility bullets.
 * (Contains detailed logging from previous debugging)
 * @param {Object} job The job item data object.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl The element after which new content should be inserted.
 * @param {Object} baseAttrsFromPlaceholder Base paragraph attributes.
 * @param {number} index The index of this item in its list.
 * @return {GoogleAppsScript.Document.Element} The last document element inserted.
 */
function renderExperienceJob(job, bodyRef, insertAfterEl, baseAttrsFromPlaceholder, index) {
  Logger.log(`  renderExperienceJob (Standard): Index ${index}, Company: "${job.company || 'N/A'}"`);
  let currentLastInsertedElement = insertAfterEl;
  let insertionIndex;

  try {
    insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1;
  } catch (e) {
    Logger.log(`    renderExperienceJob: Error getting child index. Defaulting to end. ${e.message}`);
    insertionIndex = bodyRef.getNumChildren();
  }

  const itemLineAttrs = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);

  // Spacing between job entries
  if (index > 0) {
    const spacingPara = bodyRef.insertParagraph(insertionIndex, "");
    let spacingParaStyle = {};
    spacingParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = 6; // Main space between jobs
    spacingParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    spacingParaStyle[DocumentApp.Attribute.LINE_SPACING] = 0.8;
    spacingPara.setAttributes(spacingParaStyle);
    currentLastInsertedElement = spacingPara;
    insertionIndex++;
  }
  
  itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0) ? 0 : 1; // For first line of job content
  itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1; // Minimal after content lines
  itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;

  // Company Name (Line 1)
  const companyText = (job.company || "Unknown Company").trim();
  if (companyText) {
    const companyPara = bodyRef.insertParagraph(insertionIndex, companyText);
    let companyAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    companyAttrs[DocumentApp.Attribute.BOLD] = true;
    companyPara.setAttributes(companyAttrs);
    currentLastInsertedElement = companyPara;
    insertionIndex++;
  }
  
  // Job Title, Location (Line 2)
  let titleText = (job.jobTitle || "Untitled Role").trim();
  const locationText = (job.location || "").trim();
  let titleLocationLine = titleText;
  if (locationText) {
      titleLocationLine += (titleLocationLine ? `, ${locationText}` : locationText);
  }

  if (titleLocationLine) {
    const titleLocPara = bodyRef.insertParagraph(insertionIndex, "");
    let titleLocAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    titleLocAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to company
    titleLocPara.setAttributes(titleLocAttrs);

    titleLocPara.appendText((job.jobTitle || "").trim()); // Append non-italic title first
    if (locationText) {
        const locStartIndex = titleLocPara.getText().length;
        titleLocPara.appendText(titleText ? `, ${locationText}` : locationText);
        titleLocPara.editAsText().setItalic(locStartIndex, titleLocPara.getText().length - 1, true);
    }
    currentLastInsertedElement = titleLocPara;
    insertionIndex++;
  }
  
  // Dates (Line 3, italic)
  const dateString = ((job.startDate || "") + (job.endDate && job.endDate.toLowerCase() !== "present" ? ` – ${job.endDate}` : (job.startDate ? " – Present" : "Date N/A"))).trim();
  if (dateString !== "Date N/A") {
    const datePara = bodyRef.insertParagraph(insertionIndex, dateString);
    let dateAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    dateAttrs[DocumentApp.Attribute.ITALIC] = true;
    dateAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to title/location
    datePara.setAttributes(dateAttrs);
    currentLastInsertedElement = datePara;
    insertionIndex++;
  }

  // Responsibilities (Bullets)
  const responsibilities = job.responsibilities || [];
  if (responsibilities.length > 0) {
    const bulletParagraphStyle = copyAttributesPreservingEnums(itemLineAttrs);
    bulletParagraphStyle[DocumentApp.Attribute.INDENT_START] = 36;
    bulletParagraphStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18;
    bulletParagraphStyle[DocumentApp.Attribute.SPACING_BEFORE] = 3; // A bit more space before first bullet
    bulletParagraphStyle[DocumentApp.Attribute.SPACING_AFTER] = 2;
    // bulletParagraphStyle[DocumentApp.Attribute.LINE_SPACING] = 1.0; // Optionally tighter for bullets

    responsibilities.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = (bulletText || "").toString().trim();
      if (trimmedBullet) {
        const bulletListItem = bodyRef.insertListItem(insertionIndex, trimmedBullet);
        if(bulletIdx === 0) { // Apply base bullet style
            bulletListItem.setAttributes(bulletParagraphStyle);
        } else { // Subsequent bullets might have tighter SPACING_BEFORE
            let subsequentBulletStyle = copyAttributesPreservingEnums(bulletParagraphStyle);
            subsequentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1;
            bulletListItem.setAttributes(subsequentBulletStyle);
        }
        bulletListItem.setGlyphType(DocumentApp.GlyphType.BULLET);
        currentLastInsertedElement = bulletListItem;
        insertionIndex++;
      }
    });
  }
  
  Logger.log(`  renderExperienceJob (Standard): Finished for "${job.company || 'N/A'}"`);
  return currentLastInsertedElement;
}


/**
 * Renders a single project item into the document body.
 * @param {Object} project The project item data object.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl Element after which to insert.
 * @param {Object} baseAttrsFromPlaceholder Base paragraph attributes.
 * @param {number} index The index of this item in its list.
 * @return {GoogleAppsScript.Document.Element} The last document element inserted.
 */
function renderProjectItem(project, bodyRef, insertAfterEl, baseAttrsFromPlaceholder, index) {
  Logger.log(`  renderProjectItem (Standard): Index ${index}, Project: "${project.projectName || 'N/A'}"`);
  let currentLastInsertedElement = insertAfterEl;
  let insertionIndex;

  try {
    insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1;
  } catch (e) {
    insertionIndex = bodyRef.getNumChildren();
  }

  const itemLineAttrs = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);

  if (index > 0) {
    const spacingPara = bodyRef.insertParagraph(insertionIndex, "");
    let spacingParaStyle = {};
    spacingParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = 6;
    spacingParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    spacingParaStyle[DocumentApp.Attribute.LINE_SPACING] = 0.8;
    spacingPara.setAttributes(spacingParaStyle);
    currentLastInsertedElement = spacingPara;
    insertionIndex++;
  }

  itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0) ? 0 : 1;
  itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.15;


  // Line 1: Project Name (Bold)
  const projNameText = (project.projectName || "Untitled Project").trim();
  if (projNameText) {
    const namePara = bodyRef.insertParagraph(insertionIndex, projNameText);
    let nameAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    nameAttrs[DocumentApp.Attribute.BOLD] = true;
    namePara.setAttributes(nameAttrs);
    currentLastInsertedElement = namePara;
    insertionIndex++;
  }

  // Line 2: Role and/or Organization
  let roleOrgLine = "";
  const roleText = (project.role || "").trim();
  const orgText = (project.organization || "").trim();
  if (roleText) roleOrgLine += roleText;
  if (orgText && orgText.toLowerCase() !== projNameText.toLowerCase()) {
    roleOrgLine += (roleOrgLine ? ", " : "") + orgText;
  }
  if (roleOrgLine) {
    const roleOrgPara = bodyRef.insertParagraph(insertionIndex, roleOrgLine);
    let roleOrgAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    roleOrgAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to project name
    // roleOrgAttrs[DocumentApp.Attribute.ITALIC] = true; // Optionally italicize
    roleOrgPara.setAttributes(roleOrgAttrs);
    currentLastInsertedElement = roleOrgPara;
    insertionIndex++;
  }

  // Line 3: Dates (Italic)
  const dateTextProj = ((project.startDate || "") + (project.endDate && project.endDate.toLowerCase() !== "present" ? ` – ${project.endDate}` : (project.startDate ? " – Present" : ""))).trim();
  if (dateTextProj) {
    const datePara = bodyRef.insertParagraph(insertionIndex, dateTextProj);
    let dateAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    dateAttrs[DocumentApp.Attribute.ITALIC] = true;
    dateAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to role/org line
    datePara.setAttributes(dateAttrs);
    currentLastInsertedElement = datePara;
    insertionIndex++;
  }

  // Description Bullets
  const descBullets = project.descriptionBullets || [];
  if (descBullets.length > 0) {
    const bulletStyle = copyAttributesPreservingEnums(itemLineAttrs);
    bulletStyle[DocumentApp.Attribute.INDENT_START] = 36;
    bulletStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18;
    bulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 3; // Space before first bullet
    bulletStyle[DocumentApp.Attribute.SPACING_AFTER] = 2;

    descBullets.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = (bulletText || "").toString().trim();
      if (trimmedBullet) {
        const li = bodyRef.insertListItem(insertionIndex, trimmedBullet);
         if(bulletIdx === 0) {
            li.setAttributes(bulletStyle);
        } else { 
            let subsequentBulletStyle = copyAttributesPreservingEnums(bulletStyle);
            subsequentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1;
            li.setAttributes(subsequentBulletStyle);
        }
        li.setGlyphType(DocumentApp.GlyphType.BULLET);
        currentLastInsertedElement = li;
        insertionIndex++;
      }
    });
  }

  // Technologies
  const technologies = project.technologies || [];
  const validTech = technologies.filter(t => typeof t === 'string' && t.trim());
  if (validTech.length > 0) {
    const techPara = bodyRef.insertParagraph(insertionIndex, "");
    let techAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    techAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (descBullets.length > 0 || roleOrgLine || dateTextProj) ? 3 : 1;
    techPara.setAttributes(techAttrs);
    techPara.appendText("Technologies: ").setBold(true).setItalic(true);
    techPara.appendText(validTech.join(', '));
    currentLastInsertedElement = techPara;
    insertionIndex++;
  }

  // GitHub Links
  const githubLinks = project.githubLinks || [];
  if (githubLinks.length > 0) {
    githubLinks.forEach((link, linkIdx) => {
      const linkName = (link.name || "Repository").trim();
      const linkUrl = (link.url || "").trim();
      if (linkUrl) {
        const ghPara = bodyRef.insertParagraph(insertionIndex, "");
        let linkAttrs = copyAttributesPreservingEnums(itemLineAttrs);
        linkAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (linkIdx === 0 && (validTech.length > 0 || descBullets.length > 0)) ? 3 : 1;
        ghPara.setAttributes(linkAttrs);
        ghPara.appendText(`GitHub (${linkName}): `);
        ghPara.appendText(linkUrl).setLinkUrl(linkUrl).setUnderline(true).setForegroundColor("#0000FF");
        currentLastInsertedElement = ghPara;
        if (linkIdx < githubLinks.length - 1) insertionIndex++;
      }
    });
  }
  Logger.log(`  renderProjectItem (Standard): Finished for "${project.projectName || 'N/A'}"`);
  return currentLastInsertedElement;
}

/**
 * Renders the Technical Skills & Certificates section.
 * Iterates through subsections and lists skills/certs.
 * @param {Object} skillsSectionData The entire data object for the "TECHNICAL SKILLS & CERTIFICATES" section.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl Element after which to insert.
 * @param {Object} baseAttrsFromPlaceholder Base paragraph attributes.
 * @return {GoogleAppsScript.Document.Element} The last document element inserted.
 */
// In DocumentService.gs
// REPLACE any existing renderTechnicalSkillsList function with this one.
// This function expects skillsSectionData to be an object:
// { title: "...", subsections: [ { name: "Category", items: [ {skill, details} OR {name, issuer, date} ] } ] }
// where the subsections and items have already been filtered by Stage 3 to include only selected skills/certs.

/**
 * Renders the Technical Skills & Certificates section using standard DocumentApp formatting.
 * It iterates through subsections (categories) and lists the selected skills/certs within each.
 *
 * @param {Object} skillsSectionData The data object for the "TECHNICAL SKILLS & CERTIFICATES" section.
 *        This object is expected to already be filtered by Stage 3 to contain only
 *        user-selected and score-approved skills/certs, grouped into their original subsections.
 *        Expected structure: { title: "...", subsections: [ { name: "CategoryName", items: [ {skill, details} or {name, issuer, issueDate} ] } ] }
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl The document element after which new content should be inserted.
 * @param {Object} baseAttrsFromPlaceholder The base paragraph attributes inherited from the section's placeholder.
 * @return {GoogleAppsScript.Document.Element} The last document element inserted for this section.
 */
function renderTechnicalSkillsList(skillsSectionData, bodyRef, insertAfterEl, baseAttrsFromPlaceholder) {
  Logger.log(`  renderTechnicalSkillsList (Standard): Processing data: ${JSON.stringify(skillsSectionData).substring(0, 250)}...`);
  let currentLastInsertedElement = insertAfterEl;
  let insertionIndex;

  try {
    insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1;
  } catch (e) {
    Logger.log(`    renderTechnicalSkillsList: Error getting child index. Defaulting to end. ${e.message}`);
    insertionIndex = bodyRef.getNumChildren();
  }

  const subsections = skillsSectionData.subsections || []; // These are now ONLY selected items, grouped by category

  if (subsections.length > 0) {
    subsections.forEach((subsection, subIndex) => {
      const categoryName = (subsection.name || "Uncategorized Skills").trim();
      const itemsInCategory = subsection.items || []; // These are ONLY selected items for this category

      if (categoryName && Array.isArray(itemsInCategory) && itemsInCategory.length > 0) {
        const itemLineAttrs = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);

        // Spacing for the first category line vs. subsequent category lines
        if (subIndex === 0) { // First skill category line in the section
          itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0; // Tight to section header (if placeholder's para spacing was handled)
        } else { // Gap between different skill categories
          itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] || 4; // Small gap (e.g., 4pt)
        }
        itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] || 2;
        itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;
        
        const skillCategoryLinePara = bodyRef.insertParagraph(insertionIndex, "");
        skillCategoryLinePara.setAttributes(itemLineAttrs);
        
        // Category Name (e.g., "Programming Languages:")
        skillCategoryLinePara.appendText(categoryName + ": ").setBold(true);
        
        let skillsCertsTextArray = [];
        itemsInCategory.forEach(item => {
          if (item.skill && item.skill.trim()) { // It's a skill item
            let skillEntry = item.skill.trim();
            if (item.details && item.details.trim()) {
              skillEntry += ` ${item.details.trim()}`; // Often details are in parentheses already from sheet
            }
            skillsCertsTextArray.push(skillEntry);
          } else if (item.name && item.name.trim()) { // It's a certificate/license item
            let certEntry = item.name.trim();
            let extras = [];
            if (item.issuer && item.issuer.trim()) extras.push(item.issuer.trim());
            if (item.issueDate && item.issueDate.trim()) extras.push(item.issueDate.trim());
            if (extras.length > 0) certEntry += ` (${extras.join(", ")})`;
            skillsCertsTextArray.push(certEntry);
          } else if (typeof item === 'string' && item.trim()) { // Fallback if item is just a string (less ideal)
            skillsCertsTextArray.push(item.trim());
          }
        });
        
        const skillsCertsString = skillsCertsTextArray.join(", ").trim();
        if (skillsCertsString) {
          skillCategoryLinePara.appendText(skillsCertsString).setBold(false).setItalic(false); // Ensure actual items are not bold unless specified by baseAttrs
        }
        
        currentLastInsertedElement = skillCategoryLinePara;
        insertionIndex++;
      }
    });
  } else {
    Logger.log("  renderTechnicalSkillsList: No selected subsections or items to render for Technical Skills.");
  }
  
  Logger.log(`  renderTechnicalSkillsList (Standard): Finished processing.`);
  return currentLastInsertedElement;
}

/**
 * Renders a single leadership/involvement item into the document body.
 * @param {Object} item The leadership item data object from finalTailoredResumeObject.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl The element after which new content should be inserted.
 * @param {Object} baseAttrsFromPlaceholder The base paragraph attributes inherited from the placeholder.
 * @param {number} index The index of this item in its list (0 for the first item).
 * @return {GoogleAppsScript.Document.Element} The last document element inserted for this item.
 */
function renderLeadershipItem(item, bodyRef, insertAfterEl, baseAttrsFromPlaceholder, index) {
  Logger.log(`  renderLeadershipItem (Standard): Index ${index}, Organization: "${item.organization || 'N/A'}"`);
  let currentLastInsertedElement = insertAfterEl;
  let insertionIndex;

  try {
    insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1;
  } catch (e) {
    insertionIndex = bodyRef.getNumChildren();
  }

  const itemLineAttrs = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);

  if (index > 0) {
    const spacingPara = bodyRef.insertParagraph(insertionIndex, "");
    let spacingParaStyle = {};
    spacingParaStyle[DocumentApp.Attribute.SPACING_BEFORE] = 6;
    spacingParaStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    spacingParaStyle[DocumentApp.Attribute.LINE_SPACING] = 0.8;
    spacingPara.setAttributes(spacingParaStyle);
    currentLastInsertedElement = spacingPara;
    insertionIndex++;
  }

  itemLineAttrs[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0) ? 0 : 1;
  itemLineAttrs[DocumentApp.Attribute.SPACING_AFTER] = 1;
  itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] = itemLineAttrs[DocumentApp.Attribute.LINE_SPACING] || 1.0;

  // Line 1: Organization (Bold)
  const orgText = (item.organization || "Unnamed Organization/Activity").trim();
  if (orgText) {
    const orgPara = bodyRef.insertParagraph(insertionIndex, orgText);
    let orgParaAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    orgParaAttrs[DocumentApp.Attribute.BOLD] = true;
    orgPara.setAttributes(orgParaAttrs);
    currentLastInsertedElement = orgPara;
    insertionIndex++;
  }

  // Line 2: Dates (Italic)
  const dateText = ((item.startDate || "") + (item.endDate && item.endDate.toLowerCase() !== "present" ? ` – ${item.endDate}` : (item.startDate ? " – Present" : ""))).trim();
  if (dateText) {
    const datePara = bodyRef.insertParagraph(insertionIndex, dateText);
    let dateAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    dateAttrs[DocumentApp.Attribute.ITALIC] = true;
    dateAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    datePara.setAttributes(dateAttrs);
    currentLastInsertedElement = datePara;
    insertionIndex++;
  }

  // Line 3: Role [, Location (Italic)]
  let roleLineText = (item.role || "").trim();
  const locationText = (item.location || "").trim();
  if (locationText) {
    roleLineText += (roleLineText ? ", " : "") + locationText;
  }
  if (roleLineText) {
    const rolePara = bodyRef.insertParagraph(insertionIndex, "");
    let roleParaAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    roleParaAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1; 
    rolePara.setAttributes(roleParaAttrs);
    rolePara.appendText((item.role || "").trim());
    if (locationText) {
        const locStartIndex = rolePara.getText().length;
        rolePara.appendText((item.role && item.role.trim() ? ", " : "") + locationText);
        rolePara.editAsText().setItalic(locStartIndex, rolePara.getText().length -1, true);
    }
    currentLastInsertedElement = rolePara;
    insertionIndex++;
  }

  // Line 4: Description (Optional)
  const descriptionText = (item.description || "").trim();
  if (descriptionText) {
    const descPara = bodyRef.insertParagraph(insertionIndex, descriptionText);
    let descParaAttrs = copyAttributesPreservingEnums(itemLineAttrs);
    descParaAttrs[DocumentApp.Attribute.SPACING_BEFORE] = 1;
    if (item.responsibilities && item.responsibilities.length > 0) {
        // descParaAttrs[DocumentApp.Attribute.INDENT_START] = 18; // Optional small indent
    }
    descPara.setAttributes(descParaAttrs);
    currentLastInsertedElement = descPara;
    insertionIndex++;
  }
  
  // Responsibilities (Bullets)
  const responsibilities = item.responsibilities || [];
  if (responsibilities.length > 0) {
    const bulletStyle = copyAttributesPreservingEnums(itemLineAttrs);
    bulletStyle[DocumentApp.Attribute.INDENT_START] = 36;
    bulletStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18;
    bulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 3; // A bit more space before first bullet
    bulletStyle[DocumentApp.Attribute.SPACING_AFTER] = 2;

    responsibilities.forEach((bulletText, bulletIdx) => {
      const trimmedBullet = (bulletText || "").toString().trim();
      if (trimmedBullet) {
        const li = bodyRef.insertListItem(insertionIndex, trimmedBullet);
        if(bulletIdx === 0) {
            li.setAttributes(bulletStyle);
        } else {
            let subsequentBulletStyle = copyAttributesPreservingEnums(bulletStyle);
            subsequentBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = 1;
            li.setAttributes(subsequentBulletStyle);
        }
        li.setGlyphType(DocumentApp.GlyphType.BULLET);
        currentLastInsertedElement = li;
        insertionIndex++;
      }
    });
  }
  
  Logger.log(`  renderLeadershipItem (Standard): Finished for "${item.organization || 'N/A'}"`);
  return currentLastInsertedElement;
}

/**
 * Renders a single honor/award item as a bullet point.
 * @param {Object} award The award item data object.
 * @param {GoogleAppsScript.Document.Body} bodyRef Reference to the document body.
 * @param {GoogleAppsScript.Document.Element} insertAfterEl Element after which to insert.
 * @param {Object} baseAttrsFromPlaceholder Base paragraph attributes.
 * @param {number} index The index of this item (used for spacing, though less critical here).
 * @return {GoogleAppsScript.Document.Element} The last document element inserted.
 */
function renderHonorItem(award, bodyRef, insertAfterEl, baseAttrsFromPlaceholder, index) { // index renamed from index_unused
    let currentLastInsertedElement = insertAfterEl; 
    let insertionIndex;
    try { insertionIndex = bodyRef.getChildIndex(currentLastInsertedElement) + 1; } 
    catch (e) { insertionIndex = bodyRef.getNumChildren(); }

    let awardText = (award.awardName || "Unnamed Award").trim();
    const detailsText = (award.details || "").trim();
    const dateText = (award.date || "").trim();

    if (detailsText) awardText += ` (${detailsText})`;
    if (dateText) awardText += (awardText.includes(detailsText) || awardText === (award.awardName || "").trim() ? ` - ` : '') + dateText;
    
    if (awardText.trim()) {
        const honorBulletStyle = copyAttributesPreservingEnums(baseAttrsFromPlaceholder);
        // For honors list, the SPACING_BEFORE on the first item is from the placeholder itself if handled correctly.
        // Or can be set to 0 here if `populateBlockPlaceholder` adjusted placeholder para's spacing.
        honorBulletStyle[DocumentApp.Attribute.SPACING_BEFORE] = (index === 0) ? 0 : 2; // Minimal space between honor items
        honorBulletStyle[DocumentApp.Attribute.SPACING_AFTER] = honorBulletStyle[DocumentApp.Attribute.SPACING_AFTER] || 2;
        honorBulletStyle[DocumentApp.Attribute.INDENT_START] = 36;
        honorBulletStyle[DocumentApp.Attribute.INDENT_FIRST_LINE] = 18;

        const honorListItem = bodyRef.insertListItem(insertionIndex, awardText.trim());
        honorListItem.setAttributes(honorBulletStyle);
        honorListItem.setGlyphType(DocumentApp.GlyphType.BULLET);
        currentLastInsertedElement = honorListItem;
        // insertionIndex++; // No increment needed if it's the only thing this function adds per item
    }
    return currentLastInsertedElement;
}


// --- DOCUMENT CLEANUP ---

/**
 * Aggressively removes all visually blank paragraphs from the document body.
 * This version aims to remove paragraphs that only contain whitespace characters (spaces, tabs, newlines).
 * @param {GoogleAppsScript.Document.Body} body The document body.
 */
function cleanupAllEmptyLines(body) {
    Logger.log("--- AGGRESSIVE cleanupAllEmptyLines START ---");
    let emptyParagraphsRemoved = 0;

    for (let i = body.getNumChildren() - 1; i >= 0; i--) {
        const element = body.getChild(i);
        if (element.getType() === DocumentApp.ElementType.PARAGRAPH) {
            const paragraph = element.asParagraph();
            let isVisuallyEmpty = false;

            if (paragraph.getNumChildren() === 0) {
                isVisuallyEmpty = true;
                // Logger.log_DEBUG(`  Para at child idx ${i} has getNumChildren()===0. Marked for removal.`);
            } else {
                let combinedText = "";
                let hasNonTextElements = false;
                for (let j = 0; j < paragraph.getNumChildren(); j++) {
                    const childInlineElement = paragraph.getChild(j);
                    if (childInlineElement.getType() === DocumentApp.ElementType.TEXT) {
                        combinedText += childInlineElement.asText().getText();
                    } else {
                        hasNonTextElements = true; break;
                    }
                }
                if (!hasNonTextElements && combinedText.replace(/\s/g, "") === "") {
                    isVisuallyEmpty = true;
                    // Logger.log_DEBUG(`  Para at child idx ${i} has only whitespace. Marked for removal.`);
                }
            }
            if (isVisuallyEmpty) {
                try {
                    paragraph.removeFromParent(); emptyParagraphsRemoved++;
                    // Logger.log_DEBUG(`    Successfully removed para at original child idx ${i}.`);
                } catch (e) { Logger.log(`    ERROR removing para at idx ${i}: ${e.toString()}`);}
            }
        }
    }
    Logger.log(`--- AGGRESSIVE cleanupAllEmptyLines END. Removed ${emptyParagraphsRemoved} visually blank paragraphs. ---`);
}


// --- MAIN DOCUMENT CREATION FUNCTION ---

/**
 * Creates a formatted resume Google Document from a resume data object.
 * It copies a template, replaces placeholders, and populates sections with data.
 *
 * @param {Object} resumeDataObject The structured resume data object (e.g., from Main.gs Stage 3).
 * @param {string} documentTitle The desired title for the new Google Document.
 * @return {string|null} The URL of the newly created Google Document, or null on error.
 */
function createFormattedResumeDoc(resumeDataObject, documentTitle) {
  let currentTask = "Function Entry & Validation"; // For error tracking
  Logger.log(`DocSvc (Standard Formatting): Starting 'createFormattedResumeDoc' for title: "${documentTitle}"`);

  if (!resumeDataObject || typeof resumeDataObject !== 'object' || !resumeDataObject.personalInfo) {
    Logger.log("ERROR (DocSvc Standard): Invalid or missing resumeDataObject or personalInfo. Cannot create document.");
    return null;
  }
  
  const pi = resumeDataObject.personalInfo;
  let finalDocTitle = "Generated Resume"; // Default
  if (documentTitle && typeof documentTitle === 'string' && documentTitle.trim() !== "") {
    finalDocTitle = documentTitle.trim();
  } else if (pi && pi.fullName) {
    finalDocTitle = `Resume - ${pi.fullName}`;
  }
  // Add a timestamp to prevent Drive naming conflicts and for versioning
  if (!finalDocTitle.match(/\(\s*Generated\s*\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
    finalDocTitle += ` (Generated ${new Date().toISOString()})`;
  }
  
  Logger.log(`DocSvc (Standard): Effective document title will be: "${finalDocTitle}"`);
  // Logger.log_DEBUG(`DocSvc (Standard): Using RESUME_TEMPLATE_DOC_ID (from Constants.gs): "${RESUME_TEMPLATE_DOC_ID}"`);

  let newDocFile; 
  let doc;      
  let body;

  try {
    // --- 1. Template Handling: Copy and Open ---
    currentTask = "Accessing and Copying Template Document";
    Logger.log(`DocSvc (Standard) - Step 1: ${currentTask} (ID from Constants.gs: RESUME_TEMPLATE_DOC_ID)`);
    const templateFile = DriveApp.getFileById(RESUME_TEMPLATE_DOC_ID); // From Constants.gs
    Logger.log(`  Template file "${templateFile.getName()}" accessed.`);
    newDocFile = templateFile.makeCopy(finalDocTitle, DriveApp.getRootFolder());
    Logger.log(`  Document copy "${newDocFile.getName()}" created (ID: ${newDocFile.getId()}).`);

    currentTask = "Opening Copied Document";
    Logger.log(`DocSvc (Standard) - Step 2: ${currentTask} (ID: ${newDocFile.getId()})`);
    doc = DocumentApp.openById(newDocFile.getId());
    Logger.log(`  Copied document "${doc.getName()}" opened for editing.`);

    body = doc.getBody();
    Logger.log("  Document body retrieved.");

    // --- 2. Populate Static Header Information (Personal Info, Summary) ---
    currentTask = "Populating Personal Info & Summary Placeholders";
    Logger.log(`DocSvc (Standard) - Step 3: ${currentTask}...`);
    if (pi) {
      findAndReplaceText_minimal(body, "{FULL_NAME}", pi.fullName);
      findAndReplaceText_minimal(body, "{CONTACT_LINE_1}", [pi.location, pi.phone, pi.email].filter(Boolean).join("  •  "));
      
      const linksPlaceholderPattern = RegExp.escape("{CONTACT_LINKS}");
      const linksRange = body.findText(linksPlaceholderPattern);
      if (linksRange) {
        const linksElement = linksRange.getElement();
        if (linksElement && linksElement.getParent() && linksElement.getParent().getType() === DocumentApp.ElementType.PARAGRAPH) {
            const linksPara = linksElement.getParent().asParagraph();
            linksPara.clear(); 
            let firstLink = true;
            const addLinkSeparator = () => { if (!firstLink) linksPara.appendText("  •  "); firstLink = false; };
            
            if (pi.linkedin)  { addLinkSeparator(); linksPara.appendText("LinkedIn").setLinkUrl(pi.linkedin).setUnderline(true).setForegroundColor("#0000FF"); }
            if (pi.portfolio) { addLinkSeparator(); linksPara.appendText("Portfolio").setLinkUrl(pi.portfolio).setUnderline(true).setForegroundColor("#0000FF"); }
            if (pi.github)    { addLinkSeparator(); linksPara.appendText("GitHub").setLinkUrl(pi.github).setUnderline(true).setForegroundColor("#0000FF"); }
        } else {
            Logger.log(`  WARN (DocSvc Standard): Parent of "{CONTACT_LINKS}" placeholder is not a Paragraph or element not found.`);
        }
      } else {
        Logger.log(`  WARN (DocSvc Standard): Placeholder "{CONTACT_LINKS}" not found in template.`);
      }
    }
    findAndReplaceText_minimal(body, "{SUMMARY_CONTENT}", resumeDataObject.summary);
    Logger.log(`  ${currentTask} finished.`);

    // --- 3. Populate Dynamic Resume Sections ---
    currentTask = "Populating Dynamic Resume Sections";
    Logger.log(`DocSvc (Standard) - Step 4: ${currentTask}...`);
    if (!resumeDataObject.sections || !Array.isArray(resumeDataObject.sections) || resumeDataObject.sections.length === 0) {
      Logger.log("  INFO (DocSvc Standard): No 'sections' array in resumeDataObject. No dynamic sections to populate.");
    } else {
      Logger.log(`  Found ${resumeDataObject.sections.length} sections in resumeDataObject. Rendering in predefined order.`);
      
      const sectionsRenderOrder = [
        "EDUCATION", "EXPERIENCE", "PROJECTS", 
        "TECHNICAL SKILLS & CERTIFICATES", 
        "LEADERSHIP & UNIVERSITY INVOLVEMENT", "HONORS & AWARDS"
      ];
      
      sectionsRenderOrder.forEach(sectionTitleToRender => {
        // Logger.log_DEBUG(`  Attempting to render section (standard): ${sectionTitleToRender}`);
        const sectionData = resumeDataObject.sections.find(s => s && s.title && s.title.toUpperCase() === sectionTitleToRender.toUpperCase());
        
        let placeholderText;
        let itemsToRenderArray = null;
        let itemRenderFunction = null; 
        let isSpecialSectionRender = false;

        switch (sectionTitleToRender) {
          case "EDUCATION":
            placeholderText = "{EDUCATION_ITEMS}";
            if (sectionData && sectionData.items && sectionData.items.length > 0) {
              itemsToRenderArray = sectionData.items; itemRenderFunction = renderEducationItem; // Use standard render
            }
            break;
          case "EXPERIENCE":
            placeholderText = "{EXPERIENCE_JOBS}";
            if (sectionData && sectionData.items && sectionData.items.length > 0) {
              itemsToRenderArray = sectionData.items; itemRenderFunction = renderExperienceJob; // Use standard render
            }
            break;
          case "PROJECTS":
            placeholderText = "{PROJECT_ITEMS}";
            if (sectionData && sectionData.subsections && sectionData.subsections.some(ss => ss.items && ss.items.length > 0)) {
              let allProjectItems = [];
              sectionData.subsections.forEach(ss => { if (ss.items) allProjectItems.push(...ss.items); });
              if (allProjectItems.length > 0) { 
                  itemsToRenderArray = allProjectItems; itemRenderFunction = renderProjectItem; // Use standard render
              }
            }
            break;
          case "TECHNICAL SKILLS & CERTIFICATES":
            placeholderText = "{TECHNICAL_SKILLS_LIST}";
            if (sectionData && sectionData.subsections && sectionData.subsections.some(ss => ss.items && ss.items.length > 0)) {
              itemsToRenderArray = [sectionData]; // Pass the whole filtered sectionData object
              itemRenderFunction = (sData, bRef, iAE, bAttrs) => renderTechnicalSkillsList(sData, bRef, iAE, bAttrs); // Standard render
              isSpecialSectionRender = true;
            }
            break;
          case "LEADERSHIP & UNIVERSITY INVOLVEMENT":
            placeholderText = "{LEADERSHIP_ITEMS}";
            if (sectionData && sectionData.items && sectionData.items.length > 0) {
              itemsToRenderArray = sectionData.items; itemRenderFunction = renderLeadershipItem; // Use standard render
            }
            break;
          case "HONORS & AWARDS":
            placeholderText = "{HONORS_LIST}";
            if (sectionData && sectionData.items && sectionData.items.length > 0) {
              itemsToRenderArray = sectionData.items; itemRenderFunction = renderHonorItem; // Standard render
            }
            break;
          default:
            Logger.log(`  WARN (DocSvc Standard): Unknown section title "${sectionTitleToRender}" in render order. Cannot map.`);
            return; 
        }

        if (itemRenderFunction && ((itemsToRenderArray && itemsToRenderArray.length > 0) || isSpecialSectionRender)) {
          Logger.log(`  Calling populateBlockPlaceholder for "${sectionTitleToRender}" with ${isSpecialSectionRender ? 'full section object' : itemsToRenderArray.length + ' items'} (using standard render function).`);
          // No documentId needed for standard render functions that don't use Docs API
          populateBlockPlaceholder(body, placeholderText, itemsToRenderArray, itemRenderFunction); 
        } else {
          Logger.log(`  INFO (DocSvc Standard): Section "${sectionTitleToRender}" data found but no items/subsections to render, or no render function. Clearing placeholder "${placeholderText}".`);
          findAndReplaceText_minimal(body, placeholderText, ""); 
        }
      }); 
    } 
    Logger.log(`  ${currentTask} finished.`);
    
    // --- 4. Final Document Cleanup ---
    currentTask = "Final Document Cleanup (Empty Lines)";
    Logger.log(`DocSvc (Standard) - Step 5: ${currentTask}...`);
    cleanupAllEmptyLines(body); 
    Logger.log("  Empty line cleanup complete.");

    // --- 5. Save and Close; Return URL ---
    currentTask = "Save and Close Document";
    Logger.log(`DocSvc (Standard) - Step 6: ${currentTask}...`);
    doc.saveAndClose();

    Logger.log(`SUCCESS (DocSvc Standard): Document "${doc.getName()}" generated. URL: ${doc.getUrl()}`);
    return doc.getUrl();

  } catch (e) {
    Logger.log(`CRITICAL ERROR (DocSvc Standard) during task "${currentTask}": ${e.message}`);
    Logger.log(`  Stack (Standard): ${e.stack}`);
    if (newDocFile && newDocFile.getId()) { 
      try { DriveApp.getFileById(newDocFile.getId()).setTrashed(true); Logger.log(`  Attempted to trash incomplete document copy (ID: ${newDocFile.getId()}).`); } 
      catch (delErr) { Logger.log(`  Error trashing incomplete copy: ${delErr.message}`);}
    }
    return null; 
  }
}
