# Microsoft Office Add-in for Administrative Document Generation

## Overview

This add-in for Microsoft Office Pro enables users to efficiently compose administrative documents in compliance with Chinese laws and regulations. By leveraging the DeepSeek API, the add-in generates high-quality content based on user-selected prompts within Word documents. It also ensures that the generated documents are properly formatted by eliminating redundant symbols, characters, and garbled codes.

## Features

- **API Integration**: Seamlessly connects to the DeepSeek API to generate content based on user input.
- **Document Formatting**: Automatically formats generated documents, removing unwanted characters and ensuring clarity.
- **Compliance**: Generates documents that adhere to Chinese laws and regulations.
- **User-Friendly**: Simple to use interface that enhances productivity for administrative tasks.

## Advantages

1. **Efficiency**: Significantly reduces the time required to draft administrative documents by automating content generation.
2. **Accuracy**: Ensures that the generated documents comply with relevant laws and regulations, minimizing legal risks.
3. **Quality**: Produces high-quality content that is clear and professional, enhancing the overall presentation of documents.
4. **Customizable**: Users can tailor prompts to meet specific needs, allowing for flexibility in document creation.
5. **Error Reduction**: Minimizes human errors by automating the formatting and content generation processes.
6. **Easy Integration**: Works directly within Microsoft Word, making it accessible without the need for additional software.

## Installation

1. Download the add-in from the [GitHub repository](link_to_your_repository).
2. Follow the installation instructions provided in the repository.

## Usage

1. Open Microsoft Word and select the text you want to use as a prompt.
2. Run the add-in by navigating to the add-ins menu.
3. The add-in will generate the document based on the selected text, formatted according to Chinese regulations.
4. Review and edit the generated document as necessary.

   <img width="922" alt="1" src="https://github.com/user-attachments/assets/32a4c9bc-fa16-46e8-bf46-d17b45a5bbf9" />

<img width="1290" alt="3" src="https://github.com/user-attachments/assets/57665326-eab1-4b99-be70-1f585c8bfddb" />


## Code Example

Here is a brief overview of the main functions used in the add-in:

```vb
Function CallDeepSeekAPI(api_key As String, inputText As String) As String
    ' Function to call DeepSeek API and return the response
    ' ...
End Function

Sub 智慧办公()
    ' Main subroutine to handle user input, API calls, and document formatting
    ' ...
End Sub
