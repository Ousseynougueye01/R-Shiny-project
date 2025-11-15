This repository contains the RShiny application for the Health Insurance Pricing Platform. The application allows users to upload an Excel file with beneficiary information, apply the pricing rules based on age brackets, territoriality, and coverage rate, and download an Excel file with the calculated premiums.

Before running the application, you must create a folder named ww in the root directory of the project. This folder must contain two files: logo.png and favicon.png. These files are used for the visual identity of the platform.

To run the application, install the required R packages and use shiny::runApp() from the project directory.

The platform includes features such as Excel import/export, automatic premium calculation, and a user-friendly interface designed for health insurance pricing.
