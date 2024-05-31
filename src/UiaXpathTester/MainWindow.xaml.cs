using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using System.Xml.XPath;

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

        /// <summary>
        /// Event handler for the "Test XPath" button click.
        /// </summary>
        /// <param name="sender">The event sender.</param>
        /// <param name="e">The event arguments.</param>
        private async void BtnTestXpath_Click(object sender, RoutedEventArgs e)
        {
            // Execute the test XPath logic asynchronously
            await Task.Run(() =>
            {
                // Update UI to show "Working..."
                Dispatcher.BeginInvoke(() =>
                {
                    LblStatus.Visibility = Visibility.Visible;
                    DtaElementData.Visibility = Visibility.Hidden;
                    LblStatus.Content = "Working...";
                });

                // Perform the actual logic
                PublishDataGrid(mainWindow: this);

                // Update UI to revert changes
                Dispatcher.BeginInvoke(() =>
                {
                    LblStatus.Visibility = Visibility.Hidden;
                    DtaElementData.Visibility = Visibility.Visible;
                });
            });
        }

        /// <summary>
        /// Publishes data from a DataGrid using the specified XPath in the provided MainWindow.
        /// </summary>
        /// <param name="mainWindow">The MainWindow instance where the DataGrid and XPath TextBox are located.</param>
        private static void PublishDataGrid(MainWindow mainWindow)
        {
            // Use the dispatcher to execute the operation on the UI thread.
            mainWindow.Dispatcher.BeginInvoke(() =>
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
            return FindElement(automationElement, xpath).Element.UIAutomationElement;
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

        #region *** Find Element ***
        // Finds an Element based on the specified xpath and UI Automation element.
        private static (int Status, Element Element) FindElement(IUIAutomationElement applicationRoot, string xpath)
        {
            // Create a new UI Automation instance.
            var session = new CUIAutomation8();

            // Split the xpath into segments based on the '|' delimiter.
            var segments = xpath.Split("|", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

            // Return status 400 if segments are null or empty.
            if (segments == null || segments.Length == 0)
            {
                return (400, default);
            }

            // Check if the xpath contains '/DOM/' for Document Object Model (DOM) searches.
            if (xpath.Contains("/DOM/", StringComparison.OrdinalIgnoreCase))
            {
                return FindElement(session, applicationRoot, xpath);
            }

            // Iterate through each segment in the xpath.
            foreach (var segment in segments)
            {
                // Find the element based on the current segment in the xpath.
                var (statusCode, element) = FindElement(xpath: segment, applicationRoot);

                // Return if the element is successfully found.
                if (statusCode == 200)
                {
                    return (statusCode, element);
                }
            }

            // Return 404 if the element is not found in any segment.
            return (404, default);
        }

        // Finds an Element based on the specified xpath and UI Automation session.
        private static (int Status, Element Element) FindElement(CUIAutomation8 session, IUIAutomationElement applicationRoot, string xpath)
        {
            // Get locator hierarchy and determine if the xpath starts from the root.
            var (isRoot, hierarchy) = GetLocatorHierarchy(xpath);

            // Return status 400 if the xpath does not specify criteria.
            if (!hierarchy.Any())
            {
                return (400, default);
            }

            // Filter out empty segments and "dom" from the hierarchy.
            var segments = hierarchy
                .Where(i => !string.IsNullOrEmpty(i) && !i.Equals("dom", StringComparison.OrdinalIgnoreCase))
                .Select(i => $"/{i}")
                .ToArray();

            // Separate the UI element segment and the element segment.
            var uiElementSegment = segments[0];
            var elementSegment = string.Join(string.Empty, segments.Skip(1));

            // Create a new UI Automation instance.
            var automation = new CUIAutomation8();

            // Determine the root element based on isRoot flag and applicationRoot availability.
            var rootElement = !isRoot && applicationRoot != null ? applicationRoot : automation.GetRootElement();

            // Find the UI Automation element based on the UI element segment in the xpath.
            var (statusCode, element) = FindElement(applicationRoot, xpath: isRoot ? $"/root{uiElementSegment}" : uiElementSegment);

            // Return 404 if the UI Automation element is not found.
            if (statusCode == 404)
            {
                return (statusCode, element);
            }

            // If there's only one segment, convert and return the UI Automation element.
            if (segments.Length == 1)
            {
                return (200, ConvertToElement(element.UIAutomationElement));
            }

            // Get the element from the Document Object Model based on the element segment.
            var (status, automationElement) = GetElementFromDocumentObjectModel(element, elementSegment);

            // Return 404 if the element is not found in the Document Object Model.
            if (automationElement == default)
            {
                return (404, default);
            }

            // Return the HTTP status code and the successfully retrieved Element.
            return (status, automationElement);
        }

        // Finds an Element based on specified criteria, such as coordinates or property-based xpath.
        private static (int Status, Element Element) FindElement(string xpath, IUIAutomationElement applicationRoot)
        {
            // Check if the application root is null and return status 404 if true.
            if (applicationRoot == null)
            {
                return (404, default);
            }

            // Try to find an element based on coordinates (cords).
            var elementByCords = GetByCords(xpath);

            // Return the result if an element is successfully found based on coordinates.
            if (elementByCords.Status == 200)
            {
                return elementByCords;
            }

            // If finding by coordinates fails, try to find an element based on property criteria.
            return GetByProperty(applicationRoot, xpath);
        }

        // Retrieves an Element from the Document Object Model (DOM) based on the specified xpath.
        private static (int Status, Element Element) GetElementFromDocumentObjectModel(Element rootElement, string xpath)
        {
            try
            {
                // Create a new UI Automation instance.
                var automation = new CUIAutomation8();

                // Get the Document Object Model (DOM) based on the rootElement's UIAutomationElement.
                var dom = new DocumentObjectModelFactory(rootElement.UIAutomationElement).NewDocumentObjectModel();

                // Extract the 'id' attribute from the specified xpath.
                var idAttribute = dom.XPathSelectElement(xpath)?.Attribute("id")?.Value;

                // Deserialize the 'id' attribute value into an array of integers.
                var id = JsonSerializer.Deserialize<int[]>(idAttribute);

                // Create a property condition based on the 'RuntimeId' property and the extracted id.
                var condition = automation.CreatePropertyCondition(UIA_PropertyIds.UIA_RuntimeIdPropertyId, id);

                // Specify the search scope in the tree.
                var treeScope = TreeScope.TreeScope_Descendants;

                // Find the UI Automation element in the DOM based on the created condition.
                rootElement.UIAutomationElement = rootElement.UIAutomationElement.FindFirst(treeScope, condition);

                // Determine the HTTP status code based on whether the element was found.
                var statusCode = rootElement.UIAutomationElement == null ? 404 : 200;

                // Convert the UI Automation element to the custom Element type.
                rootElement = rootElement.UIAutomationElement == null
                    ? default
                    : ConvertToElement(rootElement.UIAutomationElement);

                // Return status 404 if the rootElement is not found.
                if (rootElement == default)
                {
                    return (404, default);
                }

                // Return the HTTP status code and the successfully retrieved Element.
                return (statusCode, rootElement);
            }
            catch (Exception e) when (e != null)
            {
                // Return status 404 if an exception occurs during the retrieval process.
                return (404, default);
            }
        }

        // Retrieves an Element based on the specified xpath, using property-based criteria.
        private static (int Status, Element Element) GetByProperty(IUIAutomationElement applicationRoot, string xpath)
        {
            // Get locator hierarchy and determine if the xpath starts from the root.
            var (isRoot, hierarchy) = GetLocatorHierarchy(xpath);

            // Return status 400 if the xpath does not specify property criteria.
            if (!hierarchy.Any())
            {
                return (400, default);
            }

            // Use the application root or get the desktop root based on isRoot flag.
            var rootElement = !isRoot && applicationRoot != null
                ? applicationRoot
                : new CUIAutomation8().GetRootElement();

            // Get the first element in the hierarchy based on property criteria.
            var automationElement = GetElementBySegment(new CUIAutomation8(), rootElement, hierarchy.First());

            // Return status 404 if the first element is not found.
            if (automationElement == default)
            {
                return (404, default);
            }

            // Iterate through the hierarchy to get the final element based on property criteria.
            foreach (var pathSegment in hierarchy.Skip(1))
            {
                automationElement = GetElementBySegment(new CUIAutomation8(), automationElement, pathSegment);

                // Return status 404 if any segment in the hierarchy is not found.
                if (automationElement == default)
                {
                    return (404, default);
                }
            }

            // Convert the UI Automation element to the custom Element type.
            var element = ConvertToElement(automationElement);

            // Return status 200 along with the successfully retrieved Element.
            return (200, element);
        }

        // Parses the xpath expression to extract locator hierarchy and determine if it starts from the desktop.
        private static (bool FromDesktop, IEnumerable<string> Hierarchy) GetLocatorHierarchy(string xpath)
        {
            // Define regular expression options for case-insensitivity and single-line mode.
            const RegexOptions RegexOptions = RegexOptions.IgnoreCase | RegexOptions.Singleline;

            // Extract values enclosed in single quotes from the xpath.
            var values = Regex.Matches(input: xpath, pattern: @"(?<==').+?(?=')").Select(i => i.Value).ToArray();

            // Determine if the xpath starts from the desktop.
            var fromDesktop = Regex.IsMatch(input: xpath, pattern: @"^(\(+)?\/(root|dom)", RegexOptions);

            // Remove the desktop prefix from the xpath if it starts from the desktop.
            xpath = fromDesktop
                ? Regex.Replace(input: xpath, pattern: @"^(\(+)?\/(root|dom)", string.Empty, RegexOptions)
                : xpath;

            // Replace values in the xpath with tokens for easier manipulation.
            var tokens = new Dictionary<string, string>();
            for (int i = 0; i < values.Length; i++)
            {
                tokens[$"value_token_{i}"] = values[i];
                xpath = xpath.Replace(values[i], $"value_token_{i}");
            }

            // Split the xpath into segments based on '/' while preserving special characters inside square brackets.
            var hierarchy = Regex
                .Split(xpath, @"\/(?=\w+|\*)(?![^\[]*\])")
                .Where(i => !string.IsNullOrEmpty(i))
                .ToArray();

            // Adjust the hierarchy segments to ensure proper formatting.
            for (int i = 0; i < hierarchy.Length; i++)
            {
                var segment = hierarchy[i];
                if (!segment.Equals("/") && !segment.EndsWith('/'))
                {
                    continue;
                }
                hierarchy[i + 1] = $"/{hierarchy[i + 1]}";
            }

            // Filter and trim empty segments from the hierarchy.
            hierarchy = hierarchy
                .Where(i => !string.IsNullOrEmpty(i) && !i.Equals("/"))
                .Select(i => i.TrimEnd('/'))
                .ToArray();

            // Replace tokens with their original values in the hierarchy.
            for (int i = 0; i < hierarchy.Length; i++)
            {
                foreach (var token in tokens)
                {
                    hierarchy[i] = hierarchy[i].Replace(token.Key, token.Value);
                }
            }

            // Return the result as a tuple.
            return (fromDesktop, hierarchy);
        }

        // Gets a UI Automation element based on the provided path segment.
        private static IUIAutomationElement GetElementBySegment(CUIAutomation8 session, IUIAutomationElement rootElement, string pathSegment)
        {
            // Get control type and property conditions based on the path segment.
            var controlTypeCondition = GetControlTypeCondition(session, pathSegment);
            var propertyCondition = GetPropertyCondition(session, pathSegment);

            // Determine if the scope is set to descendants (true if path starts with '/').
            var isDescendants = pathSegment.StartsWith('/');
            var scope = isDescendants ? TreeScope.TreeScope_Descendants : TreeScope.TreeScope_Children;

            // Declare a variable to hold the UI Automation condition.
            IUIAutomationCondition condition;

            // Choose the appropriate condition based on control type and property conditions.
            if (controlTypeCondition == default && propertyCondition != default)
            {
                condition = propertyCondition;
            }
            else if (controlTypeCondition != default && propertyCondition == default)
            {
                condition = controlTypeCondition;
            }
            else if (controlTypeCondition != default && propertyCondition != default)
            {
                condition = session.CreateAndCondition(controlTypeCondition, propertyCondition);
            }
            else
            {
                // If no conditions are specified, return default.
                return default;
            }

            // Extract index from the path segment.
            var index = Regex.Match(input: pathSegment, pattern: @"(?<=\[)\d+(?=])").Value;
            var isIndex = int.TryParse(index, out int indexOut);

            // If no index is specified, find the first matching element.
            if (!isIndex)
            {
                return rootElement.FindFirst(scope, condition);
            }

            // If an index is specified, find all matching elements and return the one at the specified index.
            var elements = rootElement.FindAll(scope, condition);

            // Return the first element if no matching elements were found; otherwise, return the element at the specified index.
            return elements.Length == 0
                ? default
                : elements.GetElement(indexOut - 1 < 0 ? 0 : indexOut - 1);
        }

        // Gets an Element based on the provided xpath, assuming it represents coordinates.
        private static (int Status, Element Element) GetByCords(string xpath)
        {
            // Attempt to get an Element with a ClickablePoint based on the provided xpath.
            var element = GetFlatPointElement(xpath);

            // If the Element could not be created (xpath does not represent coordinates), return status 404.
            if (element == null)
            {
                return (404, default);
            }

            // Return status 200 along with the successfully retrieved Element.
            return (200, element);
        }

        // Gets an Element with a ClickablePoint based on the provided xpath, assuming it represents coordinates.
        private static Element GetFlatPointElement(string xpath)
        {
            // Check if the xpath matches the expected pattern for coordinates.
            var isCords = Regex.IsMatch(input: xpath, pattern: "(?i)//cords\\[\\d+,\\d+]");

            // If the xpath does not represent coordinates, return null.
            if (!isCords)
            {
                return null;
            }

            // Deserialize the coordinates from the xpath string.
            var cords = JsonSerializer.Deserialize<int[]>(Regex.Match(input: xpath, pattern: "\\[\\d+,\\d+]").Value);

            // Create and return an Element with a ClickablePoint based on the deserialized coordinates.
            return new Element
            {
                ClickablePoint = new ClickablePoint(xpos: cords[0], ypos: cords[1])
            };
        }

        // Creates a UI Automation condition based on control type specified in a path segment.
        private static IUIAutomationCondition GetControlTypeCondition(CUIAutomation8 session, string pathSegment)
        {
            // Flags for property conditions and binding flags for reflection.
            const PropertyConditionFlags ConditionFlags = PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            const BindingFlags BindingFlags = BindingFlags.Public | BindingFlags.Static;
            const StringComparison Compare = StringComparison.OrdinalIgnoreCase;

            // Local function to get the UI Automation control type ID based on the control type name.
            static int GetControlTypeId(string controlTypeName)
            {
                // Binding flags for reflection.
                const StringComparison Compare = StringComparison.OrdinalIgnoreCase;

                // Get the fields representing UI Automation control type IDs.
                var fields = typeof(UIA_ControlTypeIds).GetFields(BindingFlags);

                // Find the field corresponding to the specified control type name.
                var id = fields
                    .FirstOrDefault(i => i.Name.Equals($"UIA_{controlTypeName}ControlTypeId", Compare))?
                    .GetValue(null);

                // Return -1 if the control type ID is not found; otherwise, cast and return the control type ID.
                return id == default ? -1 : (int)id;
            }

            // Ensure the path segment ends with brackets to avoid issues in regex parsing.
            pathSegment = pathSegment.LastIndexOf('[') == -1 ? $"{pathSegment}[]" : pathSegment;

            // Extract the control type name from the path segment using regex.
            var typeSegment = Regex.Match(input: pathSegment, pattern: @"(?<=((\/)+)?)\w+(?=\)?\[)").Value;

            // Determine condition flags based on control type segment (partial match or exact match).
            var conditionFlag = typeSegment.StartsWith("partial", Compare)
                ? ConditionFlags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring
                : ConditionFlags;

            // Remove "partial" from the control type segment.
            typeSegment = typeSegment.Replace("partial", string.Empty, Compare);

            // Get the UI Automation control type ID for the control type segment.
            var controlTypeId = GetControlTypeId(typeSegment);

            // If the control type name is empty or the ID is not found, return default.
            if (string.IsNullOrEmpty(typeSegment) || controlTypeId == -1)
            {
                return default;
            }

            // Create a property condition based on the control type ID and condition flags.
            return session
                .CreatePropertyConditionEx(UIA_PropertyIds.UIA_ControlTypePropertyId, controlTypeId, conditionFlag);
        }

        // Creates a UI Automation condition based on property segments specified in a path.
        private static IUIAutomationCondition GetPropertyCondition(CUIAutomation8 session, string pathSegment)
        {
            // Flags for property conditions and binding flags for reflection.
            const PropertyConditionFlags ConditionFlags = PropertyConditionFlags.PropertyConditionFlags_IgnoreCase;
            const BindingFlags BindingFlags = BindingFlags.Public | BindingFlags.Static;
            const StringComparison Compare = StringComparison.OrdinalIgnoreCase;

            // Local function to get the UI Automation property ID based on the property name.
            static int GetPropertyId(string propertyName, BindingFlags bindingFlags)
            {
                // Binding flags for reflection.
                const StringComparison Compare = StringComparison.OrdinalIgnoreCase;

                // Get the fields representing UI Automation property IDs.
                var fields = typeof(UIA_PropertyIds).GetFields(bindingFlags);

                // Find the field corresponding to the specified property name.
                var id = fields
                    .FirstOrDefault(i => i.Name.Equals($"UIA_{propertyName}PropertyId", Compare))?
                    .GetValue(null);

                // Return -1 if the property ID is not found; otherwise, cast and return the property ID.
                return id == default ? -1 : (int)id;
            }

            // List to store individual property conditions.
            var conditions = new List<IUIAutomationCondition>();

            // Split the path segment into individual segments and process each one.
            var segments = Regex.Match(pathSegment, @"(?<=\[).*(?=\])").Value.Split(" and ").Select(i => $"[{i}]");

            foreach (var segment in segments)
            {
                // Extract the type segment (e.g., @Edit, @Name) and value segment.
                var typeSegment = Regex.Match(segment, @"(?<=@)\w+").Value;

                // Determine condition flags based on type segment (partial match or exact match).
                var conditionFlag = typeSegment.StartsWith("partial", Compare)
                    ? ConditionFlags | PropertyConditionFlags.PropertyConditionFlags_MatchSubstring
                    : ConditionFlags;

                // Remove "partial" from the type segment.
                typeSegment = typeSegment.Replace("partial", string.Empty, Compare);

                // Extract the value segment enclosed in single or double quotes.
                var valueSegment = Regex.Match(input: segment, pattern: @"(?<=\[@\w+=('|"")).*(?=('|"")])").Value;

                // Get the UI Automation property ID for the type segment.
                var propertyId = GetPropertyId(typeSegment, BindingFlags);

                // Skip if the property ID is not found.
                if (propertyId == -1)
                {
                    continue;
                }

                // Create a property condition based on the property ID, value, and condition flags.
                var condition = session.CreatePropertyConditionEx(propertyId, valueSegment, conditionFlag);

                // Add the condition to the list.
                conditions.Add(condition);
            }

            // If no conditions are created, return default.
            if (conditions.Count == 0)
            {
                return default;
            }

            // If there's only one condition, return it. Otherwise, create an AND condition from the list.
            return conditions.Count == 1
                ? conditions.First()
                : session.CreateAndConditionFromArray(conditions.ToArray());
        }
        #endregion

        /// <summary>
        /// Extracts element data from an <see cref="IUIAutomationElement"/>.
        /// </summary>
        /// <param name="element">The UI Automation element to extract data from.</param>
        /// <returns>An <see cref="ObservableCollection{ElementData}"/> containing element data.</returns>
        public static ObservableCollection<ElementData> ExtractElementData(this IUIAutomationElement element)
        {
            // Get properties starting with "Current" from the UI Automation element type
            var properties = DocumentObjectModelFactory.GetElementAttributes(element);

            // Create a collection of ElementData from the properties
            var collection = properties.Select(i => new ElementData { Property = i.Key, Value = i.Value });

            // Return the element data collection as an ObservableCollection
            return new ObservableCollection<ElementData>(collection);
        }

        // Converts an IUIAutomationElement to an Element.
        private static Element ConvertToElement(IUIAutomationElement automationElement)
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
    }

    /// <summary>
    /// Factory class for creating Document Object Models (DOM) based on UI Automation elements.
    /// </summary>
    /// <param name="rootElement">The root UI Automation element to be used for creating document object models.</param>
    public class DocumentObjectModelFactory(IUIAutomationElement rootElement)
    {
        private readonly IUIAutomationElement _rootElement = rootElement;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentObjectModelFactory"/> class.
        /// Uses the root UI Automation element obtained from <see cref="CUIAutomation8.GetRootElement"/>.
        /// </summary>
        public DocumentObjectModelFactory()
            : this(new CUIAutomation8().GetRootElement())
        { }

        /// <summary>
        /// Gets the attributes of a UI Automation element using default timeout of 5 seconds.
        /// </summary>
        /// <param name="element">The UI Automation element.</param>
        /// <returns>
        /// A dictionary containing the attributes of the UI Automation element.
        /// The keys are attribute names, and the values are their corresponding values.
        /// </returns>
        public static IDictionary<string, string> GetElementAttributes(IUIAutomationElement element)
        {
            // Use the default timeout of 5 seconds to get element attributes.
            return FormatElementAttributes(element, TimeSpan.FromSeconds(5));
        }

        /// <summary>
        /// Gets the attributes of a UI Automation element using the specified timeout.
        /// </summary>
        /// <param name="element">The UI Automation element.</param>
        /// <param name="timeout">The maximum time to wait for obtaining element attributes.</param>
        /// <returns>
        /// A dictionary containing the attributes of the UI Automation element.
        /// The keys are attribute names, and the values are their corresponding values.
        /// </returns>
        public static IDictionary<string, string> GetElementAttributes(IUIAutomationElement element, TimeSpan timeout)
        {
            // Get element attributes using the specified timeout.
            return FormatElementAttributes(element, timeout);
        }

        /// <summary>
        /// Creates a new XML document object model (XDocument) based on the root UI Automation element and its descendants.
        /// </summary>
        /// <returns>An XDocument representing the document object model of the root UI Automation element and its descendants.</returns>
        public XDocument NewDocumentObjectModel()
        {
            // Create a new instance of CUIAutomation8 for UI Automation operations.
            var automation = new CUIAutomation8();

            // Use the root element if available, otherwise, obtain the root element.
            var element = _rootElement ?? automation.GetRootElement();

            // Call the core method to create the document object model.
            return NewDocumentObjectModel(automation, element);
        }

        /// <summary>
        /// Creates a new XML document object model (XDocument) based on the provided UI Automation element and its descendants.
        /// </summary>
        /// <param name="element">The UI Automation element to start building the document from.</param>
        /// <returns>An XDocument representing the document object model of the provided UI Automation element and its descendants.</returns>
        public static XDocument NewDocumentObjectModel(IUIAutomationElement element)
        {
            // Create a new instance of CUIAutomation8 for UI Automation operations.
            var automation = new CUIAutomation8();

            // Call the core method to create the document object model.
            return NewDocumentObjectModel(automation, element);
        }

        // Creates a new XML document object model (XDocument) based on the UI Automation element and its descendants.
        private static XDocument NewDocumentObjectModel(CUIAutomation8 automation, IUIAutomationElement element)
        {
            // Read the document object model into a list of XML strings.
            var xmlData = ReadDocumentObjectModel(automation, element);

            // Combine the XML strings into a single XML document.
            var xml = "<Root>" + string.Join("\n", xmlData) + "</Root>";

            try
            {
                // Attempt to parse the XML string into an XDocument.
                return XDocument.Parse(xml);
            }
            catch (Exception e) when (e != null)
            {
                // If an exception occurs during parsing, create an XDocument with an error message.
                return XDocument.Parse($"<Root><Error>{e.GetBaseException().Message}</Error></Root>");
            }
        }

        // Reads the Document Object Model (DOM) of a UI Automation element and returns it as a list of XML strings.
        private static List<string> ReadDocumentObjectModel(CUIAutomation8 automation, IUIAutomationElement element)
        {
            // Initialize a list to store XML representations of the DOM.
            var xml = new List<string>();

            // Get the tag name and attributes of the root UI Automation element.
            var tagName = GetTagName(element, TimeSpan.FromSeconds(5));
            var attributes = ExtractElementAttributes(element);

            // Add the opening tag of the root element to the XML list.
            xml.Add($"<{tagName} {attributes}>");

            // Create a condition to select all elements.
            var condition = automation.CreateTrueCondition();
            var treeWalker = automation.CreateTreeWalker(condition);

            // Get the first child element of the root element.
            var childElement = treeWalker.GetFirstChildElement(element);

            // Recursively read the DOM starting from the first child element.
            while (childElement != null)
            {
                // Read the DOM of the child element and add its XML representation to the list.
                var nodeXml = ReadDocumentObjectModel(automation, childElement);
                xml.AddRange(nodeXml);

                // Move to the next sibling element.
                childElement = treeWalker.GetNextSiblingElement(childElement);
            }

            // Add the closing tag of the root element to the XML list.
            xml.Add($"</{tagName}>");

            // Return the list of XML representations of the DOM.
            return xml;
        }

        // Retrieves attributes of a UI Automation element.
        private static string ExtractElementAttributes(IUIAutomationElement element)
        {
            // Get the attributes of the UI Automation element.
            var attributes = FormatElementAttributes(element, TimeSpan.FromSeconds(5));

            // Get the runtime ID of the UI Automation element and serialize it to a JSON string.
            var runtime = element.GetRuntimeId().OfType<int>();
            var id = JsonSerializer.Serialize(runtime);
            attributes.Add("id", id);

            // Create a list to store XML node representations of the non-empty and non-whitespace attributes.
            var xmlNode = new List<string>();

            // Iterate through each attribute and filter out empty and whitespace-only keys and values.
            foreach (var item in attributes)
            {
                // Skip attributes with empty or whitespace-only keys.
                if (string.IsNullOrEmpty(item.Key) || Regex.IsMatch(input: item.Key, pattern: "$\\s+^"))
                {
                    continue;
                }

                // Skip attributes with empty values.
                if (string.IsNullOrEmpty(item.Value))
                {
                    continue;
                }

                // Add the attribute to the XML node list.
                xmlNode.Add($"{item.Key}=\"{item.Value}\"");
            }

            // Join the XML node representations into a single string and return it.
            return string.Join(" ", xmlNode);
        }

        // Formats the attributes of a UI Automation element and returns them as a dictionary.
        private static Dictionary<string, string> FormatElementAttributes(IUIAutomationElement element, TimeSpan timeout)
        {
            // Formats a string by replacing special XML characters.
            static string Format(string input)
            {
                // Check if the input string is null or empty.
                if (string.IsNullOrEmpty(input))
                {
                    // Return an empty string if the input is null or empty.
                    return string.Empty;
                }

                // Replace special XML characters in the input string.
                // - "&" is replaced with "&amp;"
                // - "\"" is replaced with "&quot;"
                // - "<" is replaced with "&lt;"
                // - ">" is replaced with "&gt;"
                return input
                    .Replace("&", "&amp;")
                    .Replace("\"", "&quot;")
                    .Replace("<", "&lt;")
                    .Replace(">", "&gt;");
            }

            // Local function to get a dictionary of formatted attributes from a UI Automation element.
            static Dictionary<string, string> Get(IUIAutomationElement element) => new()
            {
                ["AcceleratorKey"] = Format(element.CurrentAcceleratorKey),
                ["AccessKey"] = Format(element.CurrentAccessKey),
                ["AriaProperties"] = Format(element.CurrentAriaProperties),
                ["AriaRole"] = Format(element.CurrentAriaRole),
                ["AutomationId"] = Format(element.CurrentAutomationId),
                ["Bottom"] = $"{element.CurrentBoundingRectangle.bottom}",
                ["Left"] = $"{element.CurrentBoundingRectangle.left}",
                ["Right"] = $"{element.CurrentBoundingRectangle.right}",
                ["Top"] = $"{element.CurrentBoundingRectangle.top}",
                ["ClassName"] = Format(element.CurrentClassName),
                ["FrameworkId"] = Format(element.CurrentFrameworkId),
                ["HelpText"] = Format(element.CurrentHelpText),
                ["IsContentElement"] = element.CurrentIsContentElement == 1 ? "true" : "false",
                ["IsControlElement"] = element.CurrentIsControlElement == 1 ? "true" : "false",
                ["IsEnabled"] = element.CurrentIsEnabled == 1 ? "true" : "false",
                ["IsKeyboardFocusable"] = element.CurrentIsKeyboardFocusable == 1 ? "true" : "false",
                ["IsPassword"] = element.CurrentIsPassword == 1 ? "true" : "false",
                ["IsRequiredForForm"] = element.CurrentIsRequiredForForm == 1 ? "true" : "false",
                ["ItemStatus"] = Format(element.CurrentItemStatus),
                ["ItemType"] = Format(element.CurrentItemType),
                ["Name"] = Format(element.CurrentName),
                ["NativeWindowHandle"] = $"{element.CurrentNativeWindowHandle}",
                ["Orientation"] = $"{element.CurrentOrientation}",
                ["ProcessId"] = $"{element.CurrentProcessId}"
            };

            // Calculate the end time by adding the specified timeout to the current time.
            var end = DateTime.Now.Add(timeout);

            // Continue trying to format attributes until the specified timeout is reached.
            while (DateTime.Now < end)
            {
                try
                {
                    // Attempt to get and return the formatted attributes.
                    return Get(element);
                }
                catch (COMException e) when (e != null)
                {
                    // Handle COMException if needed. Currently, the exception is caught and the loop continues.
                }
            }

            // Return an empty dictionary if attribute formatting is not successful within the specified timeout.
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        // Retrieves the tag name of a UI Automation element within a specified timeout.
        private static string GetTagName(IUIAutomationElement element, TimeSpan timeout)
        {
            // Calculate the expiration time based on the current time and the specified timeout.
            var expires = DateTime.Now.Add(timeout);

            // Continue trying to retrieve the tag name until the expiration time is reached.
            while (DateTime.Now < expires)
            {
                try
                {
                    // Get the control type name using reflection based on the UI Automation element's CurrentControlType.
                    var controlType = typeof(UIA_ControlTypeIds).GetFields()
                        .Where(f => f.FieldType == typeof(int))
                        .FirstOrDefault(f => (int)f.GetValue(null) == element.CurrentControlType)?.Name;

                    // Use a regular expression to extract the tag name from the control type name.
                    return Regex
                        .Match(input: controlType, pattern: "(?i)//cords\\[\\d+,\\d+]", RegexOptions.None)
                        .Value;
                }
                catch (COMException e) when (e != null)
                {
                    // Handle COMException if needed. Currently, the exception is caught and the loop continues.
                }
            }

            // Return an empty string if the tag name is not found within the specified timeout.
            return string.Empty;
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
