# UIA XPath Tester v1.0

[![Build & Release](https://github.com/g4-api/uia-xpath-tester/actions/workflows/GithubActions.yml/badge.svg?branch=main)](https://github.com/g4-api/uia-xpath-tester/actions/workflows/GithubActions.yml)  

## Overview

The UIA XPath Tester is a simple WPF application designed to assist users in testing UI Automation (UIA) XPath expressions. With this tool, you can input XPath expressions, initiate tests, and view the results.

## Features

- **XPath Testing:** Input your UIA XPath expressions for testing.
- **Test Button:** Click the "▶ _Test" button to initiate the test operation.
- **Result Display:** View the results in a DataGrid (hidden by default).
- **Status Feedback:** The application provides status feedback through a labeled area.

## Getting Started

Follow these steps to use the UIA XPath Tester:

1. **Clone the Repository:**
   ```
   git clone https://github.com/your-username/UIAXpathTester.git
   ```

2. **Open in Visual Studio:**
   - Open the solution file (`UIAXpathTester.sln`) in Visual Studio.

3. **Build the Solution:**
   - Build the solution to ensure all dependencies are resolved.

4. **Run the Application:**
   - Start the application using the "Start Debugging" (F5) or "Start without Debugging" (Ctrl+F5) option.

5. **Input XPath Expression:**
   - Enter your UIA XPath expression in the provided TextBox.

6. **Initiate Test:**
   - Click the "▶ _Test" button to initiate the test.

7. **View Results:**
   - Results will be displayed in the DataGrid.

## Additional Information

- **MainWindow.xaml:** Contains the XAML markup for the UI layout.
- **MainWindow.xaml.cs:** Contains the code-behind file with event handlers and logic.

## Dependencies

- This application is built using WPF (Windows Presentation Foundation).

## Screenshots

![Application Screenshot](/screenshots/1.png)

## Contributing

Contributions are welcome! If you find any issues or have improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
