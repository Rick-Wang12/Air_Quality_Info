
class FileSharer:

    '''
    Upload the report constructed by the app, and let users can download it by the new_filelink.url
    '''

    def __init__(self, filepath, api_key):
        self.filepath = filepath
        self.api_key = api_key

    def share(self):
        from filestack import Client
        client = Client(self.api_key) # Add API key
        new_filelink = client.upload(filepath = self.filepath)

        return new_filelink.url # Click the url and check the file uploaded in default browser
