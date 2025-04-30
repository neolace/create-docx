# Create-Docx

## Description

Create-Docx is a tool for generating `.docx` documents programmatically. It simplifies the process of creating Word
documents with custom content and formatting.

## Installation

To install the dependencies, run:

```bash
npm install
```

## Usage

Import the library and use it to generate `.docx` files:

```javascript
const createDocx = require('create-docx');

// Example usage
createDocx.generate({
  title: 'Sample Document',
  content: 'This is a sample document generated programmatically.',
});
```

## License

This project is licensed under the MIT License. See the LICENSE file for details.
