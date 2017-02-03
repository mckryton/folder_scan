@d-folder_scan
Feature: set scan folder

  As a user I want to select a folder containing files, so that I can get a list
  of those files.

  Scenario: use existing folder name
    Given folder name is saved in XL list file
    When I run the folder scan
    Then this folder name is taken

  Scenario: ask for folder name
    Given no folder name is set in the XL list file
    When I run the folder scan
    Then a choose folder dialog pops up
      And the script will use the selected folder from the dialog

  Scenario: cancel choose folder dialog
    Given no folder name is set in the XL list file
    When I cancel the choose folder dialog
    Then the script stops
