// ==UserScript==
// @name         Add UL Button to Product Links
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Adds a small "UL" button next to links containing "productId" in their href.
// @author       Your Name
// @match       https://iq.ulprospector.com/en/profile*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    function addUlButtons() {
        const links = document.querySelectorAll('a');
        const baseUrl = "https://www.ulsestandards.org/uls-standardsdocs/standarddocuments.aspx?CatalogDepartmentId=1&Standard=";
        const buttonStyles = 'margin-left: 5px; padding: 2px 5px; background-color: #007bff; color: white; text-decoration: none; border-radius: 3px; font-size: 0.8em; cursor: pointer; display: inline-block; vertical-align: middle;';

        for (const link of links) {
            if (link.href && link.href.includes('productId=')) {
                try {
                    const url = new URL(link.href);
                    const productId = url.searchParams.get('productId');

                    if (productId) {
                        // Remove "UL" prefix if it exists
                        const modifiedProductId = productId.startsWith('UL') ? productId.substring(2) : productId;

                        // Construct the new URL
                        const newHref = baseUrl + modifiedProductId;

                        // Check if a button already exists to avoid duplicates
                        if (link.nextElementSibling && link.nextElementSibling.classList.contains('ul-product-button')) {
                            continue; // Skip if button is already there
                        }

                        // Create the new button element
                        const button = document.createElement('a');
                        button.href = newHref;
                        button.textContent = 'UL'; // Or could use an icon/emoji
                        button.className = 'ul-product-button'; // Add a class for styling
                        button.style.cssText = buttonStyles; // Assign the styles

                        // Insert the button after the link
                        if (link.nextSibling) {
                            link.parentNode.insertBefore(button, link.nextSibling);
                        } else {
                            link.parentNode.appendChild(button);
                        }
                    }
                } catch (e) {
                    console.error("Error processing link href:", link.href, e);
                }
            }
        }
    }

    // Run the function when the document is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', addUlButtons);
    } else {
        addUlButtons();
    }

    // Optional: If content is dynamically loaded, you might need MutationObserver
    // Example:
    // const observer = new MutationObserver(addUlButtons);
    // observer.observe(document.body, { childList: true, subtree: true });

})();
