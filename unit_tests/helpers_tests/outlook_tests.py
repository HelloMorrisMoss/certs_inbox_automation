import unittest
from unittest.mock import MagicMock, patch

from helpers.outlook_helpers import add_categories_to_mail, remove_categories_from_mail


class TestAddCategoriesToMail(unittest.TestCase):

    @patch('win32com.client.Dispatch')
    def test_add_single_category(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch()
        mail.Categories = ''
        category = 'red'

        # Act
        add_categories_to_mail(mail, category)

        # Assert
        expected_categories = 'Red Category'
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_add_multiple_categories(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch()
        mail.Categories = ''
        categories = ["blue", "red"]

        # Act
        add_categories_to_mail(mail, categories)

        # Assert
        expected_categories = "Blue Category, Red Category"
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_add_single_category_with_existing_category(self, mock_dispatch):
        # Arrange
        mail = MagicMock()
        mail.Categories = 'Red Category'
        category = 'blue'

        # Act
        add_categories_to_mail(mail, category)

        # Assert
        expected_categories = 'Red Category, Blue Category'
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_invalid_categories_type(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch()
        category = '123'

        # Act/Assert
        with self.assertRaises(ValueError):
            add_categories_to_mail(mail, category)

    @patch('win32com.client.Dispatch')
    def test_invalid_categories_list_type(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch()
        categories = ["blue", "invalid_color"]

        # Act/Assert
        with self.assertRaises(ValueError):
            add_categories_to_mail(mail, categories)


class TestRemoveCategoriesFromMail(unittest.TestCase):

    @patch('win32com.client.Dispatch')
    def test_remove_single_category(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch
        mail.Categories = 'Red Category'
        category = 'red'

        # Act
        remove_categories_from_mail(mail, category)

        # Assert
        expected_categories = ''
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_remove_multiple_categories(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch
        mail.Categories = 'Blue Category, Red Category'
        categories = ["blue", "red"]

        # Act
        remove_categories_from_mail(mail, categories)

        # Assert
        expected_categories = ""
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_remove_nonexistent_category(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch
        mail.Categories = 'Red Category'
        category = 'blue'

        # Act
        remove_categories_from_mail(mail, category)

        # Assert
        expected_categories = 'Red Category'
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

    @patch('win32com.client.Dispatch')
    def test_invalid_categories_type(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch
        category = '123'

        # Act/Assert
        with self.assertRaises(ValueError):
            remove_categories_from_mail(mail, category)

    @patch('win32com.client.Dispatch')
    def test_invalid_categories_list_type(self, mock_dispatch):
        # Arrange
        mail = mock_dispatch
        categories = ["blue", "invalid_color"]

        # Act/Assert
        with self.assertRaises(ValueError):
            remove_categories_from_mail(mail, categories)

    @patch('win32com.client.Dispatch')
    def test_remove_multiple_categories(self, mock_dispatch):
        # Arrange
        mail = MagicMock()
        mail.Categories = 'Blue Category, Red Category, Green Category'
        categories = ["blue", "red"]

        # Act
        remove_categories_from_mail(mail, categories)

        # Assert
        expected_categories = "Green Category"
        self.assertEqual(mail.Categories, expected_categories)
        mail.Save.assert_called_once()

if __name__ == '__main__':
    unittest.main()
