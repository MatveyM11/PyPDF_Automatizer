(function() {
    'use strict'; // Strict mode helps catch common coding errors and "unsafe" actions.

    // Create a new script element to load the XLSX library
    const script = document.createElement('script');
    // Set the source of the script to the URL of the XLSX library
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.4/xlsx.full.min.js";
    // Define what happens once the script is successfully loaded
    script.onload = function() {
        console.log('XLSX library loaded'); // Log a message to confirm the library is loaded

        // Function to find the payment button and add an event listener to it
        function checkButton() {
            // Use querySelector to find the payment button based on its class and name attribute
            const paymentButton = document.querySelector('input.button.alt.new-css-class[name="woocommerce_checkout_place_order"]');
            // Check if the payment button was found
            if (paymentButton) {
                console.log('Payment button found:', paymentButton); // Log the found button
                // Add an event listener to the payment button to handle clicks
                paymentButton.addEventListener('click', (event) => {
                    console.log('Payment button was pressed'); // Log the button press
                    event.preventDefault(); // Prevent the default form submission behavior

                    // Collect form data by getting values from input elements by their IDs
                    const formData = {
                        billing_first_name: document.getElementById('billing_first_name')?.value || '', // Get first name
                                               billing_last_name: document.getElementById('billing_last_name')?.value || '', // Get last name
                                               billing_company: document.getElementById('billing_company')?.value || '', // Get company name
                                               billing_country: document.getElementById('billing_country')?.value || '', // Get country
                                               billing_address_1: document.getElementById('billing_address_1')?.value || '', // Get address line 1
                                               billing_address_2: document.getElementById('billing_address_2')?.value || '', // Get address line 2
                                               billing_city: document.getElementById('billing_city')?.value || '', // Get city
                                               billing_state: document.querySelector('#select2-billing_state-container')?.getAttribute('title') || '', // Get province from title attribute
                                               billing_postcode: document.getElementById('billing_postcode')?.value || '', // Get postal code
                                               billing_phone: document.getElementById('billing_phone')?.value || '', // Get phone number
                                               billing_email: document.getElementById('billing_email')?.value || '', // Get email address
                                               billing_renovacion: document.getElementById('billing_renovacion')?.value || '', // Get renewal info
                                               billing_renovados: document.getElementById('billing_renovados')?.value || '', // Get type of renewal
                                               billing_date: document.getElementById('billing_date')?.value || '' // Get date of birth
                    };

                    console.log('Collected Form Data:', formData); // Log the collected form data

                    // Create an array of arrays to represent the form data in a format suitable for Excel
                    const data = [
                        ["Campo", "Valor"], // Column headers
                        ["Nombre", formData.billing_first_name], // First name
                        ["Apellidos", formData.billing_last_name], // Last name
                        ["Empresa", formData.billing_company], // Company name
                        ["País", formData.billing_country], // Country
                        ["Dirección 1", formData.billing_address_1], // Address line 1
                        ["Dirección 2", formData.billing_address_2], // Address line 2
                        ["Ciudad", formData.billing_city], // City
                        ["Provincia", formData.billing_state], // Province
                        ["Código postal", formData.billing_postcode], // Postal code
                        ["Teléfono", formData.billing_phone], // Phone number
                        ["Correo electrónico", formData.billing_email], // Email address
                        ["Renovación", formData.billing_renovacion], // Renewal info
                        ["Tipo de renovación", formData.billing_renovados], // Type of renewal
                        ["Fecha de nacimiento", formData.billing_date] // Date of birth
                    ];

                    // Create a new workbook for Excel
                    const wb = XLSX.utils.book_new();
                    // Convert the array of arrays to a sheet and add it to the workbook
                    const ws = XLSX.utils.aoa_to_sheet(data);
                    // Append the created sheet to the workbook
                    XLSX.utils.book_append_sheet(wb, ws, "FormData");

                    // Write the workbook to a Blob object for downloading
                    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                    const blob = new Blob([wbout], { type: 'application/octet-stream' }); // Create a Blob from the workbook
                    const url = URL.createObjectURL(blob); // Create a URL for the Blob
                    const link = document.createElement('a'); // Create an anchor element
                    link.href = url; // Set the href to the Blob URL
                    link.download = 'form_data.xlsx'; // Set the download attribute to specify the filename
                    document.body.appendChild(link); // Append the link to the document
                    link.click(); // Programmatically click the link to trigger the download
                    document.body.removeChild(link); // Remove the link from the document

                    console.log('Form data saved in Excel format'); // Log that the form data was saved

                    event.target.closest('form').submit(); // Submit the form after saving the data
                });
                clearInterval(checkTimer); // Clear the interval to stop checking for the button
            } else {
                console.log('Payment button not found, retrying...'); // Log that the button was not found and retry
            }
        }

        // Set an interval to repeatedly check for the payment button every 1.5 seconds
        var checkTimer = setInterval(checkButton, 1500);
    };

    // Append the script element to the document head to start loading the XLSX library
    document.head.appendChild(script);
})();

