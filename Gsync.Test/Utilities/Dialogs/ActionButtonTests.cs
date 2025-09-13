using Gsync.Utilities.Dialogs;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Gsync.Test.Utilities.Dialogs
{
    [TestClass]
    public class ActionButtonTests
    {
        [TestMethod]
        public void Constructor_WithButtonAndAction_InitializesCorrectly()
        {
            // Arrange
            var button = new Button();
            var action = new Action(() => { });

            // Act
            var actionButton = new ActionButton(button, DialogResult.OK, action);

            // Assert
            Assert.AreEqual(button, actionButton.Button);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
        }

        [TestMethod]
        public void Constructor_WithNameButtonTextAndAction_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });

            // Act
            var actionButton = new ActionButton(name, buttonText, action);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
        }

        [TestMethod]
        public void Constructor_WithNameButtonTextActionAndTemplate_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });
            var template = new Button { BackColor = Color.Red };

            // Act
            var actionButton = new ActionButton(name, buttonText, action, template);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(Color.Red, actionButton.Button.BackColor);
        }

        [TestMethod]
        public void Constructor_WithNameButtonTextDialogResultAndAction_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });

            // Act
            var actionButton = new ActionButton(name, buttonText, DialogResult.OK, action);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
        }

        [TestMethod]
        public void Constructor_WithNameButtonTextDialogResultActionAndTemplate_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });
            var template = new Button { BackColor = Color.Red };

            // Act
            var actionButton = new ActionButton(name, buttonText, DialogResult.OK, action, template);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
            Assert.AreEqual(Color.Red, actionButton.Button.BackColor);
        }

        [TestMethod]
        public void Constructor_WithNameButtonImageButtonTextDialogResultAndAction_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });
            var image = new Bitmap(10, 10);

            // Act
            var actionButton = new ActionButton(name, image, buttonText, DialogResult.OK, action);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
            Assert.AreEqual(image, actionButton.Button.Image);
        }

        [TestMethod]
        public void Constructor_WithNameButtonImageButtonTextDialogResultActionAndTemplate_InitializesCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var action = new Action(() => { });
            var image = new Bitmap(10, 10);
            var template = new Button { BackColor = Color.Red };

            // Act
            var actionButton = new ActionButton(name, image, buttonText, DialogResult.OK, action, template);

            // Assert
            Assert.AreEqual(name, actionButton.Name);
            Assert.AreEqual(buttonText, actionButton.Button.Text);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
            Assert.AreEqual(image, actionButton.Button.Image);
            Assert.AreEqual(Color.Red, actionButton.Button.BackColor);
        }

        [TestMethod]
        public void FromButton_CreatesActionButtonCorrectly()
        {
            // Arrange
            var button = new Button();
            var action = new Action(() => { });

            // Act
            var actionButton = ActionButton.FromButton(button, DialogResult.OK, action);

            // Assert
            Assert.AreEqual(button, actionButton.Button);
            Assert.AreEqual(action, actionButton.Delegate);
            Assert.AreEqual(DialogResult.OK, actionButton.Button.DialogResult);
        }

        [TestMethod]
        public void MakeButton_CreatesButtonCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var actionButton = new ActionButton(name, buttonText, new Action(() => { }));

            // Act
            var button = actionButton.MakeButton(buttonText);

            // Assert
            Assert.AreEqual(buttonText, button.Text);
            Assert.AreEqual(name, button.Name);
        }

        [TestMethod]
        public void MakeButton_WithImage_CreatesButtonCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var image = new Bitmap(10, 10);
            var actionButton = new ActionButton(name, buttonText, new Action(() => { }));

            // Act
            var button = actionButton.MakeButton(buttonText, image);

            // Assert
            Assert.AreEqual(buttonText, button.Text);
            Assert.AreEqual(name, button.Name);
            Assert.AreEqual(image, button.Image);
        }

        [TestMethod]
        public void MakeButton_WithImageAndDialogResult_CreatesButtonCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var image = new Bitmap(10, 10);
            var actionButton = new ActionButton(name, buttonText, new Action(() => { }));

            // Act
            var button = actionButton.MakeButton(buttonText, image, DialogResult.OK);

            // Assert
            Assert.AreEqual(buttonText, button.Text);
            Assert.AreEqual(name, button.Name);
            Assert.AreEqual(image, button.Image);
            Assert.AreEqual(DialogResult.OK, button.DialogResult);
        }

        [TestMethod]
        public void MakeButton_WithDialogResult_CreatesButtonCorrectly()
        {
            // Arrange
            string name = "TestButton";
            string buttonText = "Click Me";
            var actionButton = new ActionButton(name, buttonText, new Action(() => { }));

            // Act
            var button = actionButton.MakeButton(buttonText, DialogResult.OK);

            // Assert
            Assert.AreEqual(buttonText, button.Text);
            Assert.AreEqual(name, button.Name);
            Assert.AreEqual(DialogResult.OK, button.DialogResult);
        }

        [TestMethod]
        public void Button_Click_InvokesAction()
        {
            // Arrange
            bool actionInvoked = false;
            var action = new Action(() => { actionInvoked = true; });
            var button = new Button();
            var actionButton = new ActionButton(button, DialogResult.OK, action);

            // Act
            button.PerformClick();

            // Assert
            Assert.IsTrue(actionInvoked);
        }
    }
}
