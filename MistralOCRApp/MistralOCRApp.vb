Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Diagnostics
Imports System.Linq

Module MistralOCRApp
    ' Directory where output files (text, markdown, images, etc.) will be saved
    ' This directory is placed on the user's Desktop.
    Private outputDir As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "MistralOCR_Output")

    ' Application for OCR using Mistral AI API
    ' Based on documentation: https://docs.mistral.ai/capabilities/document/

    Async Function GetFileInfo(apiKey As String, fileId As String) As Task(Of String)
        Try
            Using client As New HttpClient()
                ' Set the authorization header
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}")

                ' Send the GET request
                Console.WriteLine($"Getting file info for ID: {fileId}...")
                Dim response As HttpResponseMessage = Await client.GetAsync($"https://api.mistral.ai/v1/files/{fileId}")

                ' Check for errors
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent = Await response.Content.ReadAsStringAsync()
                    Throw New Exception($"Failed to get file info with status code {response.StatusCode}: {errorContent}")
                End If

                ' Read and return the response content
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As Exception
            Console.WriteLine($"An error occurred while getting file info: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Sub Main()
        ' Make sure the output directory exists
        If Not Directory.Exists(outputDir) Then
            Directory.CreateDirectory(outputDir)
        End If

        ' Prompt user for API key and document path
        Console.Write("Enter your Mistral API key: ")
        Dim apiKey As String = Console.ReadLine().Trim()

        Console.Write("Select document source (1 for local file, 2 for URL, 3 for existing file ID): ")
        Dim sourceOption As String = Console.ReadLine().Trim()

        Dim documentSource As String = ""

        If sourceOption = "1" Then
            Console.Write("Enter the path to your document file: ")
            documentSource = Console.ReadLine().Trim()
            If Not File.Exists(documentSource) Then
                Console.WriteLine("File does not exist.")
                Return
            End If
            ProcessLocalDocument(apiKey, documentSource)
        ElseIf sourceOption = "2" Then
            Console.Write("Enter the URL to your document file: ")
            documentSource = Console.ReadLine().Trim()
            ProcessUrlDocument(apiKey, documentSource)
        ElseIf sourceOption = "3" Then
            Console.Write("Enter the file ID: ")
            Dim fileId As String = Console.ReadLine().Trim()
            If String.IsNullOrWhiteSpace(fileId) Then
                Console.WriteLine("File ID cannot be empty.")
                Return
            End If
            ProcessFileId(apiKey, fileId)
        Else
            Console.WriteLine("Invalid option selected.")
            Return
        End If
    End Sub

    Sub ProcessLocalDocument(apiKey As String, filePath As String)
        Try
            Console.WriteLine("Uploading file to Mistral API...")

            ' Upload the file first (blocking call using .Result for simplicity)
            Dim fileId As String = UploadFile(apiKey, filePath).Result

            If String.IsNullOrEmpty(fileId) Then
                Console.WriteLine("Failed to upload file.")
                Return
            End If

            Console.WriteLine($"File uploaded successfully with ID: {fileId}")

            ' Wait a moment for the server to process the uploaded file
            Console.WriteLine("Waiting for the server to process the uploaded file...")
            System.Threading.Thread.Sleep(2000) ' Wait 2 seconds

            ' Process using the file ID method now
            ProcessFileId(apiKey, fileId)
        Catch ex As Exception
            Console.WriteLine($"An error occurred when processing local file: {ex.Message}")
            If ex.InnerException IsNot Nothing Then
                Console.WriteLine($"Inner exception: {ex.InnerException.Message}")
                Console.WriteLine($"Stack trace: {ex.StackTrace}")
            End If
        End Try
    End Sub

    Async Function UploadFile(apiKey As String, filePath As String) As Task(Of String)
        Try
            Using client As New HttpClient()
                ' Set the authorization header
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}")

                ' Create the multipart form data content
                Using content As New MultipartFormDataContent()
                    ' Add the purpose parameter
                    content.Add(New StringContent("ocr"), "purpose")

                    ' Add the file
                    Dim fileContent As New ByteArrayContent(File.ReadAllBytes(filePath))
                    content.Add(fileContent, "file", Path.GetFileName(filePath))

                    ' Send the POST request
                    Console.WriteLine("Uploading file...")
                    Dim response As HttpResponseMessage = Await client.PostAsync("https://api.mistral.ai/v1/files", content)

                    ' Check for errors
                    If Not response.IsSuccessStatusCode Then
                        Dim errorContent = Await response.Content.ReadAsStringAsync()
                        Throw New Exception($"File upload failed with status code {response.StatusCode}: {errorContent}")
                    End If

                    ' Read the response content
                    Dim responseString As String = Await response.Content.ReadAsStringAsync()

                    ' Parse the JSON response to get the file ID
                    Dim json As JObject = JObject.Parse(responseString)
                    Return json("id").ToString()
                End Using
            End Using
        Catch ex As Exception
            Console.WriteLine($"An error occurred during file upload: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Sub ProcessUrlDocument(apiKey As String, documentUrl As String)
        Try
            ' Create the JSON payload
            Dim jsonPayload As String = $"{{""model"": ""mistral-ocr-latest"", ""document"": {{""type"": ""document_url"", ""document_url"": ""{documentUrl}""}}, ""include_image_base64"": true}}"

            ' Process the document
            ProcessDocument(apiKey, jsonPayload)
        Catch ex As Exception
            Console.WriteLine($"An error occurred when processing URL: {ex.Message}")
        End Try
    End Sub

    Sub ProcessFileId(apiKey As String, fileId As String)
        Try
            ' Check if file exists and is available
            Dim fileInfo As String = GetFileInfo(apiKey, fileId).Result
            Console.WriteLine($"File info: {fileInfo}")

            ' Get a signed URL for the file
            Dim signedUrl As String = GetSignedUrl(apiKey, fileId).Result

            If String.IsNullOrEmpty(signedUrl) Then
                Console.WriteLine("Failed to get signed URL for the file.")
                Return
            End If

            Console.WriteLine($"Got signed URL for file: {signedUrl}")

            ' Create the JSON payload with the signed URL
            Dim jsonPayload As String = $"{{""model"": ""mistral-ocr-latest"", ""document"": {{""type"": ""document_url"", ""document_url"": ""{signedUrl}""}}, ""include_image_base64"": true}}"

            ' Process the document
            ProcessDocument(apiKey, jsonPayload)
        Catch ex As Exception
            Console.WriteLine($"An error occurred when processing file ID: {ex.Message}")
            If ex.InnerException IsNot Nothing Then
                Console.WriteLine($"Inner exception: {ex.InnerException.Message}")
            End If
        End Try
    End Sub

    Async Function GetSignedUrl(apiKey As String, fileId As String) As Task(Of String)
        Try
            Using client As New HttpClient()
                ' Set the authorization header
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}")

                ' Send the GET request to get a signed URL with 24 hour expiry
                Console.WriteLine($"Getting signed URL for file ID: {fileId}...")
                Dim response As HttpResponseMessage = Await client.GetAsync($"https://api.mistral.ai/v1/files/{fileId}/url?expiry=24")

                ' Check for errors
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent = Await response.Content.ReadAsStringAsync()
                    Throw New Exception($"Failed to get signed URL with status code {response.StatusCode}: {errorContent}")
                End If

                ' Read the response content
                Dim responseContent = Await response.Content.ReadAsStringAsync()
                Console.WriteLine($"URL response: {responseContent}")

                ' Parse the JSON response to get the URL
                Dim json As JObject = JObject.Parse(responseContent)
                If json("url") IsNot Nothing Then
                    Return json("url").ToString()
                Else
                    Console.WriteLine("URL not found in response")
                    Return String.Empty
                End If
            End Using
        Catch ex As Exception
            Console.WriteLine($"An error occurred while getting signed URL: {ex.Message}")
            Return String.Empty
        End Try
    End Function

    Sub ProcessDocument(apiKey As String, jsonPayload As String)
        Try
            Console.WriteLine("Processing document...")

            ' Send the OCR request and get the response (blocking call for simplicity)
            Dim responseTask As Task(Of String) = SendOcrRequest(apiKey, jsonPayload)
            If Not responseTask.Wait(TimeSpan.FromMinutes(5)) Then
                Console.WriteLine("Request timed out after 5 minutes.")
                Return
            End If

            Dim responseString As String = responseTask.Result

            ' Parse the JSON response
            Dim json As JObject = JObject.Parse(responseString)
            Console.WriteLine("JSON parsed successfully. Starting to process pages...")

            ' Log the full response structure for debugging
            Dim jsonString As String = json.ToString(Newtonsoft.Json.Formatting.None)
            Console.WriteLine("Response structure:")
            Console.WriteLine(jsonString.Substring(0, Math.Min(500, jsonString.Length)) + "...")

            ' Check if pages array exists
            If json("pages") Is Nothing Then
                Console.WriteLine("WARNING: 'pages' array not found in response")
                Console.WriteLine("Saving raw response for analysis...")

                ' Ensure output directory exists (in case not already created)
                Directory.CreateDirectory(outputDir)
                Dim debugJsonPath As String = Path.Combine(outputDir, "response.json")
                File.WriteAllText(debugJsonPath, responseString, Encoding.UTF8)
                Console.WriteLine($"Response saved to: {debugJsonPath}")
                Return
            End If

            ' Build the combined text content
            Dim textBuilder As New StringBuilder()
            Console.WriteLine($"Processing {json("pages").Count()} pages...")

            Dim pageIndex As Integer = 0
            For Each page As JObject In json("pages")
                Try
                    pageIndex += 1
                    Dim pageNumber As Integer = 0
                    If page("page_number") IsNot Nothing Then
                        pageNumber = page("page_number").ToObject(Of Integer)()
                    ElseIf page("metadata") IsNot Nothing AndAlso page("metadata")("page_number") IsNot Nothing Then
                        pageNumber = page("metadata")("page_number").ToObject(Of Integer)()
                    Else
                        pageNumber = pageIndex
                    End If

                    Dim textContent As String = ""
                    If page("markdown") IsNot Nothing Then
                        textContent = page("markdown").ToString()
                    ElseIf page("text") IsNot Nothing Then
                        textContent = page("text").ToString()
                    ElseIf page("content") IsNot Nothing Then
                        textContent = page("content").ToString()
                    End If

                    If String.IsNullOrEmpty(textContent) Then
                        Console.WriteLine($"WARNING: No text content found for page {pageNumber}")
                        Console.WriteLine($"Available properties for page {pageNumber}: {String.Join(", ", JObject.FromObject(page).Properties().Select(Function(p) p.Name))}")
                        textContent = "[No text content found]"
                    End If

                    textBuilder.AppendLine($"--- Page {pageNumber} ---")
                    textBuilder.AppendLine(textContent)
                    textBuilder.AppendLine()

                    ' Process images if available
                    Try
                        If page("images") IsNot Nothing AndAlso page("images").Count() > 0 Then
                            Console.WriteLine($"Processing {page("images").Count()} images for page {pageNumber}...")
                            Dim pageImagesDir As String = Path.Combine(outputDir, $"page_{pageNumber}_images")
                            Directory.CreateDirectory(pageImagesDir)

                            For Each image As JObject In page("images")
                                Dim imageId As String = "img_" & Guid.NewGuid().ToString()
                                If image("id") IsNot Nothing Then
                                    imageId = image("id").ToString()
                                End If

                                If image("image_base64") IsNot Nothing Then
                                    Dim base64Data As String = image("image_base64").ToString()
                                    Dim base64String As String = base64Data
                                    If base64Data.Contains(",") Then
                                        base64String = base64Data.Substring(base64Data.IndexOf(",") + 1)
                                    End If

                                    Try
                                        Dim imageBytes As Byte() = Convert.FromBase64String(base64String)
                                        Dim imagePath As String = Path.Combine(pageImagesDir, $"{imageId}.png")
                                        File.WriteAllBytes(imagePath, imageBytes)
                                        Console.WriteLine($"Saved image: {imagePath}")
                                    Catch imgEx As Exception
                                        Console.WriteLine($"Failed to save image {imageId}: {imgEx.Message}")
                                    End Try
                                End If
                            Next
                        End If
                    Catch imgEx As Exception
                        Console.WriteLine($"Error processing images for page {pageNumber}: {imgEx.Message}")
                    End Try
                Catch pageEx As Exception
                    Console.WriteLine($"Error processing page: {pageEx.Message}")
                End Try
            Next

            ' Ensure that the output directory exists before saving files
            Directory.CreateDirectory(outputDir)

            ' Save the combined text to output.txt
            Dim textPath As String = Path.Combine(outputDir, "output.txt")
            File.WriteAllText(textPath, textBuilder.ToString(), Encoding.UTF8)
            Console.WriteLine($"Text output saved to: {textPath}")

            ' Save markdown file if available
            If json("pages")(0)("markdown") IsNot Nothing Then
                Dim markdownBuilder As New StringBuilder()
                For Each page As JObject In json("pages")
                    If page("markdown") IsNot Nothing Then
                        markdownBuilder.AppendLine(page("markdown").ToString())
                        markdownBuilder.AppendLine()
                    End If
                Next
                Dim markdownPath As String = Path.Combine(outputDir, "output.md")
                File.WriteAllText(markdownPath, markdownBuilder.ToString(), Encoding.UTF8)
                Console.WriteLine($"Markdown output saved to: {markdownPath}")
            End If

            ' Save raw JSON response for debugging
            Dim jsonPath As String = Path.Combine(outputDir, "response.json")
            File.WriteAllText(jsonPath, responseString, Encoding.UTF8)
            Console.WriteLine($"Raw JSON response saved to: {jsonPath}")

            Console.WriteLine("Processing complete.")
            Console.WriteLine($"Files saved on your Desktop in the '{outputDir}' folder.")

            ' Ask if user wants to open the output folder
            Console.Write("Do you want to open the output folder? (y/n): ")
            Dim openFolder As String = Console.ReadLine().Trim().ToLower()
            If openFolder = "y" Then
                Process.Start(New ProcessStartInfo(outputDir) With {.UseShellExecute = True})
            End If

        Catch ex As Exception
            Console.WriteLine($"An error occurred during processing: {ex.Message}")
            Console.WriteLine(ex.StackTrace)
        End Try
    End Sub


    Async Function SendOcrRequest(apiKey As String, jsonPayload As String) As Task(Of String)
        Try
            Using client As New HttpClient()
                ' Set timeout to 5 minutes since OCR can take time for large documents
                client.Timeout = TimeSpan.FromMinutes(5)

                ' Set the authorization header
                client.DefaultRequestHeaders.Clear()
                client.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}")

                ' Log the request payload (for debugging, removing sensitive info as needed)
                Console.WriteLine("Request payload (partial):")
                Dim debugPayload As String = jsonPayload
                If debugPayload.Length > 500 Then
                    debugPayload = debugPayload.Substring(0, 500) + "... [truncated]"
                End If
                Console.WriteLine(debugPayload)

                ' Create the content with appropriate encoding and media type
                Dim content As New StringContent(jsonPayload, Encoding.UTF8, "application/json")
                content.Headers.ContentType = New MediaTypeHeaderValue("application/json")

                ' Send the POST request
                Console.WriteLine("Sending request to Mistral AI API...")
                Dim response As HttpResponseMessage = Await client.PostAsync("https://api.mistral.ai/v1/ocr", content)

                ' Check for errors
                If Not response.IsSuccessStatusCode Then
                    Dim errorContent = Await response.Content.ReadAsStringAsync()
                    Console.WriteLine($"Error response: {errorContent}")
                    Throw New Exception($"API request failed with status code {response.StatusCode}: {errorContent}")
                End If

                Console.WriteLine("Response received successfully.")
                Return Await response.Content.ReadAsStringAsync()
            End Using
        Catch ex As TaskCanceledException
            Throw New Exception("The request timed out. This could be due to a large document or server issues.", ex)
        Catch ex As HttpRequestException
            Throw New Exception($"HTTP request error: {ex.Message}", ex)
        Catch ex As Exception
            Throw New Exception($"Unexpected error during API request: {ex.Message}", ex)
        End Try
    End Function
End Module
