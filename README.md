# WordToHtmlAddin
This is an initial POC that includes a Word Add-in and a node.js server that receives a file from the Word add-in and convert it to html

## Running the project locally:
1. Clone the project.
2. Cd to the cloned project folder and run 
```
npm install
```
3. Start the server on your localhost with this command:
```
node server.js
```
4. Start Word and open a new or existing document.
5. Install the Word add-in using the [/manifest/GetDoc_AppManifest.xml](manifest/) file.
   Please follow [these](https://docs.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins) instructions to install the Word add-in.
6. Click on the "Submit" button.

The "Submit" button passes the current opened document in Word to the project folder in the server side and gives it the name tranFile.docx, and generates a new file with the name tranFile.html, that can be opened with any web browser.

## NOTES: 
1. This POC assumes that you are installing the Word add-in in the same machine that you are running the node.js server. To run the server on different machine or in the cloud, you must update the hardcoded URLs in the project files (You can replace "localhost:8080" everywhere in the project with the host name or IP of the server)
2. In This POC we are using http protocol, and didn't add any security layers.

