const express = require('express');
const multer = require('multer');
const docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const path = require('path');
const fs = require('fs');

// Initialize the app and set up the server
const app = express();
const port = 3000;

// Set up static files and view engine
app.use(express.static('public'));
app.set('view engine', 'ejs');
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Set up file upload destination
const upload = multer({ dest: 'uploads/' });

// Home route to render the form
app.get('/', (req, res) => {
  res.render('index');
});

// Route to handle the file upload and processing
app.post('/generate', upload.single('template'), (req, res) => {
  const { file } = req;
  const data = req.body; // Capturing the user input data

  if (!file) {
    return res.status(400).send('No file uploaded.');
  }

  try {

    // Read the uploaded DOCX file and process it with docxtemplater
    const docxPath = path.join(__dirname, file.path);
    fs.readFile(docxPath, (err, content) => {
      if (err) {
        return res.status(500).send('Error reading the uploaded file.');
      }

      const zip = new PizZip(content);
      const doc = new docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
      });
      // Render the DOCX with the data
      doc.render(data);
      const buf = doc.getZip().generate({ type: 'nodebuffer' });

      // Send the generated file to the user
      const outputPath = path.join(__dirname, 'uploads', 'generated_output.docx');
      fs.writeFileSync(outputPath, buf);

      res.download(outputPath, 'generated_output.docx', (err) => {
        if (err) {
          console.error('Error sending file:', err);
        }
        fs.unlinkSync(docxPath);  // Clean up the uploaded file
        fs.unlinkSync(outputPath); // Clean up the generated file
      });
    });
  } catch (error) {
    console.log('adasdasda', error);
    res.status(500).send('Error processing the template.');
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
