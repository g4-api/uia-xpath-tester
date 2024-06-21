using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;

using UIAutomationClient;

using UiaXpathTester.Models;

namespace UiaXpathTester
{
    /// <summary>
    /// Interaction logic for the main application window.
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
        }

        // Event handler for the "Test XPath" button click.
        private async void BtnTestXpath_Click(object sender, RoutedEventArgs e)
        {
            // Update UI to show "Working..."
            await Dispatcher.BeginInvoke(new Action(() =>
            {
                LblStatus.Visibility = Visibility.Visible;
                DtaElementData.Visibility = Visibility.Hidden;
                BtnTestXpath.IsEnabled = false;
                LblStatus.Content = "Working...";
                TxtElapsedTime.Text = "";
            }));

            // Start the stopwatch to measure elapsed time
            var stopwatch = Stopwatch.StartNew();

            // Perform the actual logic asynchronously
            await Task.Run(() => PublishDataGrid(mainWindow: this));

            // Stop the stopwatch
            stopwatch.Stop();

            // Update UI to revert changes and show elapsed time
            await Dispatcher.BeginInvoke(new Action(() =>
            {
                LblStatus.Visibility = Visibility.Hidden;
                DtaElementData.Visibility = Visibility.Visible;
                BtnTestXpath.IsEnabled = true;
                TxtElapsedTime.Text = $"Elapsed Time: {stopwatch.ElapsedMilliseconds} ms";
            }));
        }

        // Publishes data from a DataGrid using the specified XPath in the provided MainWindow.
        private static async Task PublishDataGrid(MainWindow mainWindow)
        {
            // Use the dispatcher to execute the operation on the UI thread.
            await mainWindow.Dispatcher.BeginInvoke(() =>
            {
                // Create a new UI Automation instance.
                var automation = new CUIAutomation8();

                // Get the UI Automation element based on the XPath from the TextBox in the MainWindow.
                var element = automation.GetElement(mainWindow.TxbXpath.Text);

                // Extract element data and set it as the ItemsSource for the DataGrid in the MainWindow.
                mainWindow.DtaElementData.ItemsSource = element.ExtractElementData();
            });
        }
    }

    /// <summary>
    /// Contains extension methods for working with UI Automation elements.
    /// </summary>
    public static class UiaExtensions
    {
        /// <summary>
        /// Gets a dictionary of attributes for the specified UI Automation element within the default timeout period of 5 seconds.
        /// </summary>
        /// <param name="element">The <see cref="IUIAutomationElement"/> to get the attributes of.</param>
        /// <returns>
        /// A dictionary of attribute names and values for the specified element, or an empty dictionary if the operation fails.
        /// </returns>
        public static IDictionary<string, string> GetAttributes(this IUIAutomationElement element)
        {
            // Call the overloaded GetAttributes method with a default timeout of 5 seconds.
            return GetAttributes(element, timeout: TimeSpan.FromSeconds(5));
        }

        /// <summary>
        /// Gets a dictionary of attributes for the specified UI Automation element within the given timeout period.
        /// </summary>
        /// <param name="element">The <see cref="IUIAutomationElement"/> to get the attributes of.</param>
        /// <param name="timeout">The maximum amount of time to attempt to get the attributes.</param>
        /// <returns>
        /// A dictionary of attribute names and values for the specified element, or an empty dictionary if the operation fails.
        /// </returns>
        public static IDictionary<string, string> GetAttributes(this IUIAutomationElement element, TimeSpan timeout)
        {
            // Return an empty dictionary if the element is null.
            if (element == null)
            {
                return new Dictionary<string, string>();
            }

            // Formats the input string to be XML-safe by replacing special characters with their corresponding XML entities.
            static string FormatXml(string input)
            {
                // Check if the input string is null or empty.
                if (string.IsNullOrEmpty(input))
                {
                    // Return an empty string if the input is null or empty.
                    return string.Empty;
                }

                // Replace special characters with their corresponding XML entities and return the result.
                return input
                    .Replace("&", "&amp;")     // Ampersand
                    .Replace("\"", "&quot;")   // Double quote
                    .Replace("'", "&apos;")    // Single quote
                    .Replace("<", "&lt;")      // Less than
                    .Replace(">", "&gt;")      // Greater than
                    .Replace("\n", "&#xA;")    // Newline
                    .Replace("\r", "&#xD;");   // Carriage return
            }

            // Formats the attributes of the UI Automation element into a dictionary.
            static Dictionary<string, string> FormatAttributes(IUIAutomationElement info) => new(StringComparer.OrdinalIgnoreCase)
            {
                ["AcceleratorKey"] = FormatXml(info.CurrentAcceleratorKey),
                ["AccessKey"] = FormatXml(info.CurrentAccessKey),
                ["AriaProperties"] = FormatXml(info.CurrentAriaProperties),
                ["AriaRole"] = FormatXml(info.CurrentAriaRole),
                ["AutomationId"] = FormatXml(info.CurrentAutomationId),
                ["Bottom"] = $"{info.CurrentBoundingRectangle.bottom}",
                ["ClassName"] = FormatXml(info.CurrentClassName),
                ["FrameworkId"] = FormatXml(info.CurrentFrameworkId),
                ["HelpText"] = FormatXml(info.CurrentHelpText),
                ["IsContentElement"] = info.CurrentIsContentElement == 1 ? "true" : "false",
                ["IsControlElement"] = info.CurrentIsControlElement == 1 ? "true" : "false",
                ["IsEnabled"] = info.CurrentIsEnabled == 1 ? "true" : "false",
                ["IsKeyboardFocusable"] = info.CurrentIsKeyboardFocusable == 1 ? "true" : "false",
                ["IsPassword"] = info.CurrentIsPassword == 1 ? "true" : "false",
                ["IsRequiredForForm"] = info.CurrentIsRequiredForForm == 1 ? "true" : "false",
                ["ItemStatus"] = FormatXml(info.CurrentItemStatus),
                ["ItemType"] = FormatXml(info.CurrentItemType),
                ["Left"] = $"{info.CurrentBoundingRectangle.left}",
                ["Name"] = FormatXml(info.CurrentName),
                ["NativeWindowHandle"] = $"{info.CurrentNativeWindowHandle}",
                ["Orientation"] = $"{info.CurrentOrientation}",
                ["ProcessId"] = $"{info.CurrentProcessId}",
                ["Right"] = $"{info.CurrentBoundingRectangle.right}",
                ["Top"] = $"{info.CurrentBoundingRectangle.top}"
            };

            // Calculate the expiration time for the timeout.
            var expiration = DateTime.Now.Add(timeout);

            // Attempt to get the attributes until the timeout expires.
            while (DateTime.Now < expiration)
            {
                try
                {
                    // Format and return the attributes of the element.
                    return FormatAttributes(element);
                }
                catch (COMException)
                {
                    // Ignore COM exceptions and continue attempting until the timeout expires.
                }
            }

            // Return an empty dictionary if the attributes could not be retrieved within the timeout.
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets a UI Automation element based on the specified xpath relative to the root element.
        /// </summary>
        /// <param name="automation">The UI Automation instance.</param>
        /// <param name="xpath">The xpath expression specifying the location of the element relative to the root element.</param>
        /// <returns>
        /// The UI Automation element found based on the xpath.
        /// Returns null if the xpath expression is invalid, does not specify criteria, or the element is not found.
        /// </returns>
        public static IUIAutomationElement GetElement(this CUIAutomation8 automation, string xpath)
        {
            // Get the root UI Automation element.
            var automationElement = automation.GetRootElement();

            // Use the FindElement method for finding an element based on the specified xpath.
            return FindElement(automationElement, xpath).Element?.UIAutomationElement;
        }

        /// <summary>
        /// Gets a UI Automation element based on the specified xpath relative to the current element.
        /// </summary>
        /// <param name="automationElement">The current UI Automation element.</param>
        /// <param name="xpath">The xpath expression specifying the location of the element relative to the current element.</param>
        /// <returns>
        /// The UI Automation element found based on the xpath.
        /// Returns null if the xpath expression is invalid, does not specify criteria, or the element is not found.
        /// </returns>
        public static IUIAutomationElement GetElement(this IUIAutomationElement automationElement, string xpath)
        {
            // Use the FindElement method for finding an element based on the specified xpath.
            return FindElement(automationElement, xpath).Element.UIAutomationElement;
        }

        /// <summary>
        /// Gets the tag name of the UI Automation element with a default timeout of 5 seconds.
        /// </summary>
        /// <param name="element">The <see cref="IUIAutomationElement"/> to get the tag name of.</param>
        /// <returns>The tag name of the element, or an empty string if the operation fails.</returns>
        public static string GetTagName(this IUIAutomationElement element)
        {
            // Call the overloaded GetTagName method with a default timeout of 5 seconds.
            return GetTagName(element, timeout: TimeSpan.FromSeconds(5));
        }

        /// <summary>
        /// Gets the tag name of the UI Automation element within the specified timeout.
        /// </summary>
        /// <param name="element">The <see cref="IUIAutomationElement"/> to get the tag name of.</param>
        /// <param name="timeout">The maximum amount of time to attempt to get the tag name.</param>
        /// <returns>The tag name of the element, or an empty string if the operation fails.</returns>
        public static string GetTagName(this IUIAutomationElement element, TimeSpan timeout)
        {
            // Calculate the expiration time for the timeout.
            var expires = DateTime.Now.Add(timeout);

            // Attempt to get the tag name until the timeout expires.
            while (DateTime.Now < expires)
            {
                try
                {
                    // Get the control type field name corresponding to the element's current control type.
                    var controlType = typeof(UIA_ControlTypeIds).GetFields()
                        .Where(f => f.FieldType == typeof(int))
                        .FirstOrDefault(f => (int)f.GetValue(null) == element.CurrentControlType)?.Name;

                    // Extract and return the tag name from the control type field name.
                    return Regex.Match(input: controlType, pattern: "(?<=UIA_).*(?=ControlTypeId)").Value;
                }
                catch (COMException)
                {
                    // Ignore COM exceptions and continue attempting until the timeout expires.
                }
            }

            // Return an empty string if the tag name could not be retrieved within the timeout.
            return string.Empty;
        }

        /// <summary>
        /// Extracts element data from an <see cref="IUIAutomationElement"/>.
        /// </summary>
        /// <param name="element">The UI Automation element to extract data from.</param>
        /// <returns>An <see cref="ObservableCollection{ElementData}"/> containing element data.</returns>
        public static ObservableCollection<ElementData> ExtractElementData(this IUIAutomationElement element)
        {
            // Get properties starting with "Current" from the UI Automation element type
            var attributes = element.GetAttributes();

            // Create a collection of ElementData from the properties
            var collection = attributes.Select(i => new ElementData { Property = i.Key, Value = i.Value });

            // Return the element data collection as an ObservableCollection
            return new ObservableCollection<ElementData>(collection);
        }

        // Finds a UI automation element based on the given XPath expression.
        private static (int Status, Element Element) FindElement(IUIAutomationElement applicationRoot, string xpath)
        {
            // Converts an IUIAutomationElement to an Element.
            static Element ConvertToElement(IUIAutomationElement automationElement)
            {
                // Generate a unique ID for the element based on the AutomationId, or use a new GUID if AutomationId is empty.
                var automationId = automationElement.CurrentAutomationId;
                var id = string.IsNullOrEmpty(automationId)
                    ? $"{Guid.NewGuid()}"
                    : automationElement.CurrentAutomationId;

                // Create a Location object based on the current bounding rectangle of the UI Automation element.
                var location = new Location
                {
                    Bottom = automationElement.CurrentBoundingRectangle.bottom,
                    Left = automationElement.CurrentBoundingRectangle.left,
                    Right = automationElement.CurrentBoundingRectangle.right,
                    Top = automationElement.CurrentBoundingRectangle.top
                };

                // Create a new Element object and populate its properties.
                var element = new Element
                {
                    Id = id,
                    UIAutomationElement = automationElement,
                    Location = location
                };

                // Return the created Element.
                return element;
            }

            // Convert the XPath expression to a UI Automation condition
            var condition = XpathParser.ConvertToCondition(xpath);

            // Return 400 status code if the XPath expression is invalid
            if (condition == null)
            {
                return (400, default);
            }

            // Determine the search scope based on the XPath expression
            var scope = xpath.StartsWith("//")
                ? TreeScope.TreeScope_Descendants
                : TreeScope.TreeScope_Children;

            // Find the first element that matches the condition within the specified scope
            var element = applicationRoot.FindFirst(scope, condition);

            // Return the status and element: 404 if not found, 200 if found
            return element == null
                ? (404, default)
                : (200, ConvertToElement(element));
        }
    }

    /// <summary>
    /// Factory class for creating a Document Object Model (DOM) representation of UI Automation elements.
    /// </summary>
    /// <param name="rootElement">The root UI Automation element.</param>
    public class DocumentObjectModelFactory(IUIAutomationElement rootElement)
    {
        // The root UI Automation element used as the starting point for creating the DOM.
        private readonly IUIAutomationElement _rootElement = rootElement;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentObjectModelFactory"/> class.
        /// </summary>
        public DocumentObjectModelFactory()
            : this(new CUIAutomation8().GetRootElement())
        { }

        /// <summary>
        /// Creates a new XML document representing the UI Automation element tree.
        /// </summary>
        /// <returns>A new <see cref="XDocument"/> representing the UI Automation element tree.</returns>
        public XDocument New()
        {
            // Create a new instance of the UI Automation object.
            var automation = new CUIAutomation8();

            // Use the root element if available; otherwise, get the desktop root element.
            var element = _rootElement ?? automation.GetRootElement();

            // Create the XML document from the UI Automation element tree.
            return New(automation, element, addDesktop: true);
        }

        /// <summary>
        /// Creates a new XML document representing the UI Automation element tree for the specified element.
        /// </summary>
        /// <param name="element">The UI Automation element.</param>
        /// <returns>A new <see cref="XDocument"/> representing the UI Automation element tree.</returns>
        public static XDocument New(IUIAutomationElement element)
        {
            // Create a new instance of the UI Automation object.
            var automation = new CUIAutomation8();

            // Create the XML document from the UI Automation element tree.
            return New(automation, element, addDesktop: true);
        }

        /// <summary>
        /// Creates a new XML document representing the UI Automation element tree for the specified automation object and element.
        /// </summary>
        /// <param name="automation">The UI Automation object.</param>
        /// <param name="element">The UI Automation element.</param>
        /// <returns>A new <see cref="XDocument"/> representing the UI Automation element tree.</returns>
        public static XDocument New(CUIAutomation8 automation, IUIAutomationElement element)
        {
            return New(automation, element, addDesktop: true);
        }

        /// <summary>
        /// Creates a new XML document representing the DOM structure of the given UI automation element.
        /// </summary>
        /// <param name="automation">The UI Automation instance to be used.</param>
        /// <param name="element">The UI Automation element to be processed.</param>
        /// <param name="addDesktop">Flag indicating whether to wrap the XML in a desktop tag.</param>
        /// <returns>An XDocument representing the XML structure of the UI element.</returns>
        public static XDocument New(CUIAutomation8 automation, IUIAutomationElement element, bool addDesktop)
        {
            // Register and generate XML data for the new DOM.
            var xmlData = Register(automation, element);

            // Construct the XML body with the tag name, attributes, and registered XML data.
            var xmlBody = string.Join("\n", xmlData);

            // Combine the XML data into a single XML string.
            var xml = addDesktop ? "<Desktop>" + xmlBody + "</Desktop>" : xmlBody;

            try
            {
                // Parse and return the XML document.
                return XDocument.Parse(xml);
            }
            catch (Exception e)
            {
                // Handle any parsing exceptions and return an error XML document.
                return XDocument.Parse($"<Desktop><Error>{e.GetBaseException().Message}</Error></Desktop>");
            }
        }

        // Registers and generates XML data for the new DOM.
        private static List<string> Register(CUIAutomation8 automation, IUIAutomationElement element)
        {
            // Initialize a list to store XML data.
            var xml = new List<string>();

            // Get the tag name and attributes of the element.
            var tagName = element.GetTagName();
            var attributes = GetElementAttributes(element);

            // Add the opening tag with attributes to the XML list.
            xml.Add($"<{tagName} {attributes}>");

            // Create a condition to find all child elements.
            var condition = automation.CreateTrueCondition();
            var treeWalker = automation.CreateTreeWalker(condition);
            var childElement = treeWalker.GetFirstChildElement(element);

            // Recursively process child elements.
            while (childElement != null)
            {
                var nodeXml = Register(automation, childElement);
                xml.AddRange(nodeXml);
                childElement = treeWalker.GetNextSiblingElement(childElement);
            }

            // Add the closing tag to the XML list.
            xml.Add($"</{tagName}>");

            // Return the complete XML data list.
            return xml;
        }

        // Gets the attributes of the specified UI Automation element as a string.
        private static string GetElementAttributes(IUIAutomationElement element)
        {
            // Get the attributes of the element.
            var attributes = element.GetAttributes();

            // Get the runtime ID of the element and serialize it to a JSON string.
            var runtime = element.GetRuntimeId().OfType<int>();
            var id = JsonSerializer.Serialize(runtime);
            attributes.Add("id", id);

            // Initialize a list to store attribute strings.
            var xmlNode = new List<string>();
            foreach (var item in attributes)
            {
                // Skip attributes with empty or whitespace-only keys or values.
                if (string.IsNullOrEmpty(item.Key) || string.IsNullOrEmpty(item.Value))
                {
                    continue;
                }

                // Add the attribute to the XML node list.
                xmlNode.Add($"{item.Key}=\"{item.Value}\"");
            }

            // Join the XML node representations into a single string and return it.
            return string.Join(" ", xmlNode);
        }
    }
}

namespace UiaXpathTester.Models
{
    /// <summary>
    /// Represents data for an element with a property and its corresponding value.
    /// </summary>
    public class ElementData
    {
        /// <summary>
        /// Gets or sets the property of the element.
        /// </summary>
        public string Property { get; set; }

        /// <summary>
        /// Gets or sets the value of the element.
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// Represents an element in the user interface.
    /// </summary>
    public class Element
    {
        /// <summary>
        /// Initializes a new instance of the Element class.
        /// </summary>
        public Element()
        {
            // Generate a new unique identifier using Guid.NewGuid() and assign it to the Id property.
            Id = $"{Guid.NewGuid()}";
        }

        /// <summary>
        /// Gets or sets the clickable point of the element.
        /// </summary>
        public ClickablePoint ClickablePoint { get; set; }

        /// <summary>
        /// Gets or sets the unique identifier of the element.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the location of the element on the screen.
        /// </summary>
        public Location Location { get; set; }

        /// <summary>
        /// Gets or sets the XML representation of the element.
        /// </summary>
        public XNode Node { get; set; }

        /// <summary>
        /// Gets or sets the UI Automation element associated with this element.
        /// </summary>
        public IUIAutomationElement UIAutomationElement { get; set; }
    }

    /// <summary>
    /// Represents the location (bounding box) of an element on the screen.
    /// </summary>
    public class Location
    {
        /// <summary>
        /// Gets or sets the bottom coordinate of the element.
        /// </summary>
        public int Bottom { get; set; }

        /// <summary>
        /// Gets or sets the left coordinate of the element.
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        /// Gets or sets the right coordinate of the element.
        /// </summary>
        public int Right { get; set; }

        /// <summary>
        /// Gets or sets the top coordinate of the element.
        /// </summary>
        public int Top { get; set; }
    }

    /// <summary>
    /// Represents a clickable point on the screen.
    /// </summary>
    /// <param name="xpos">The X-coordinate of the clickable point.</param>
    /// <param name="ypos">The Y-coordinate of the clickable point.</param>
    public class ClickablePoint(int xpos, int ypos)
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ClickablePoint"/> class with default coordinates (0, 0).
        /// </summary>
        public ClickablePoint()
            : this(xpos: 0, ypos: 0)
        { }

        /// <summary>
        /// Gets or sets the X-coordinate of the clickable point.
        /// </summary>
        public int XPos { get; set; } = xpos;

        /// <summary>
        /// Gets or sets the Y-coordinate of the clickable point.
        /// </summary>
        public int YPos { get; set; } = ypos;
    }
}
