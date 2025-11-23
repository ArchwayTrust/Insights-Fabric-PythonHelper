import msal
import requests
import json
import base64
import time
from notebookutils import mssparkutils

class Emails:
    """
    A class to interact with Microsoft Graph API for managing emails.
    Requires an app registration with delegated User.Read, Mail.ReadWrite, Mail.ReadWrite.Shared, offline_access and public client flow enabled.

    Attributes:
        tennant_id (str): Tenant ID for Azure authentication.
        client_id (str): Client ID for Azure authentication.
        akv_url (str): Azure Key Vault URL, in the form https://{key_vault_name}.vault.azure.net/.
        refresh_secret_name (str): Name of the secret in the Azure Key Vault that holds the refresh token.
        app (msal.PublicClientApplication): MSAL instance for Azure authentication.
        access_token (str): Token for authenticated access to the Graph API.
    """

    REDIRECT_URI = "https://login.microsoftonline.com/common/oauth2/nativeclient"
    SCOPES = ["User.Read", "Mail.ReadWrite", "Mail.Send"]

    def __init__(self, tennant_id, client_id, akv_url, refresh_secret_name):
        """
        Initializes the GraphAPI_Emails class with required Azure and Graph API parameters.

        Parameters:
            tennant_id (str): Tenant ID for Azure authentication.
            client_id (str): Client ID for Azure authentication.
            akv_url (str): Azure Key Vault URL.
            refresh_secret_name (str): Name of the secret in the Azure Key Vault.
        """
        self.tennant_id = tennant_id
        self.client_id = client_id
        self.akv_url = akv_url  # URL for Azure Key Vault
        self.refresh_secret_name = refresh_secret_name  # Refresh secret name in the Azure Key Vault
        
        # Initialize the MSAL confidential client
        self.app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tennant_id}"
        )

        self.token_response = None
        self.access_token = None
        self.refresh_token = None

        print("If you have previously authenticated run connect().")
        print("If you need to trigger a device authenication flow run get_initial_tokens().")
   
    def connect(self):
        print("Retrieving refresh token from Azure Key Vault:")
        try:
            # Get latest refresh token
            self.refresh_token = mssparkutils.credentials.getSecret(self.akv_url, self.refresh_secret_name)
            print("Success!")
        except Exception as e:
            self.token_response = None
            self.access_token = None
            print("Couldn't retrieve secret from key vault.")
            print("Ensure that the notebook user has permissions on the key vault. Then try running get_initial_tokens()")
            raise Exception(f"Issue retrieving secret from key vault. Error: {e}") from None

        print("Using stored refresh token to acquire access token and get user details:")
        try:
            # Get access token.
            self.token_response = self.app.acquire_token_by_refresh_token(self.refresh_token, self.SCOPES)
            self.access_token = self.token_response["access_token"]
            self.refresh_token = self.token_response["refresh_token"]
            self._store_refresh_token(self.refresh_token)
            self.user = self.get_user()
            self.user_principal_name = self.user.get("userPrincipalName")
            self.user_id = self.user.get("id")
            print(f"You are authenticated as {self.user_principal_name} ({self.user_id}).")

        except Exception as e:
            self.access_token = None
            print("Couldn't get access token. Try running get_initial_tokens()")
            raise Exception(f"Couldn't get access token. Error: {e}") from None

    def _store_refresh_token(self, refresh_token):
        """
        Stores the given refresh token in the Azure Key Vault.

        This method updates the Azure Key Vault secret with the new refresh token
        for future authentications.

        Parameters:
            refresh_token (str): The refresh token to be stored in the Azure Key Vault.
        """
        # Get credentials for key vault using workbook owner.
        akv_cred = mssparkutils.credentials.getToken('keyvault')

        # Prepare the URL and headers for the Azure Key Vault request
        url = f"{self.akv_url}/secrets/{self.refresh_secret_name}?api-version=7.4"
        headers = {
            "Authorization": f"Bearer {akv_cred}",
            "Content-Type": "application/json"
        }

        # Convert the refresh token to a JSON formatted string
        data = json.dumps({"value": refresh_token})

        # Make the PUT request to the Azure Key Vault to update the secret
        response = requests.put(url, headers=headers, data=data)

        # Check the response from the Azure Key Vault
        if response.status_code == 200:
            print("Refresh secret updated successfully.")
        else:
            print("Failed to update secret.")
            print("Status code:", response.status_code)
            print("Response: <response text omitted to avoid leaking sensitive information>")
            raise Exception(f"Failed to update Secret. Code: {response.status_code}")


    def get_initial_tokens(self):
        """
        Initiates the device flow to acquire initial access and refresh tokens.

        This method uses MSAL to start a device flow authentication, printing out
        a user code and waiting for the user to authenticate. If successful, it 
        stores the new access token and refresh token.
        """
        # Initiate the device flow with specified scopes
        flow = self.app.initiate_device_flow(scopes=self.SCOPES)

        # Check if the user_code is part of the flow response
        if "user_code" not in flow:
            raise ValueError(f"Failed to create device flow. Err: {json.dumps(flow, indent=4)}")

        # Set an expiration time for the flow (60 seconds from now)
        flow['expires_at'] = int(time.time()) + 60

        # Display the authentication message to the user
        print(flow["message"])

        # Attempt to acquire the token by the device flow
        result = self.app.acquire_token_by_device_flow(flow)

        # Check if the authentication was successful
        if "access_token" in result:
            print("Successfully authenticated.")
            # Store the access token and refresh token from the result
            self.access_token = result['access_token']
            self.refresh_token = result['refresh_token']
            # Store the new refresh token in the key vault
            self._store_refresh_token(self.refresh_token)
        else:
            # Handle authentication failure
            print(f"Authentication failed. Result was: {result}")
            raise Exception(f"Authentication failed. Error: {result}")

    def search_message_by_subject_and_sender(self, subject, sender_email, only_search_inbox=True, only_return_latest=True, shared_mailbox_email=None):
        """
        Searches for an email by its subject and sender's email using Microsoft Graph API with pagination.

        This method queries the Microsoft Graph API to find all emails that match
        the given subject and sender's email address, handling pagination to gather all results.

        Parameters:
            subject (str): The subject of the email to search for.
            sender_email (str): The email address of the sender of the email.
            only_search_inbox (bool): Defaults to True. Limits search to only the main inbox.
            only_return_latest (bool): Defaults to True. If True, returns the ID of the latest message.
            shared_mailbox_email (str): Optional. The email address of the shared mailbox to search in.

        Returns:
            str or list: The ID of the latest email or a list of emails that match the search criteria.
                         Returns None if no match is found.
        """
        headers = {'Authorization': f'Bearer {self.access_token}'}

        # Determine the user ID or shared mailbox email to use in the endpoint
        mailbox_id = shared_mailbox_email if shared_mailbox_email else self.user_id

        if only_search_inbox:
            endpoint = f"https://graph.microsoft.com/v1.0/users/{mailbox_id}/mailFolders/Inbox/messages"
        else:
            endpoint = f"https://graph.microsoft.com/v1.0/users/{mailbox_id}/messages"

        query_parameters = {
            '$filter': f"subject eq '{subject}' and from/emailAddress/address eq '{sender_email}'",
            '$select': 'id,subject,from,receivedDateTime,parentFolderId,hasAttachments'
        }

        messages = []
        while endpoint:
            response = requests.get(endpoint, headers=headers, params=query_parameters)
            if response.status_code != 200:
                # Handle unsuccessful API call
                raise Exception(f"Error searching message by subject: {response.json()}")

            data = response.json()
            messages.extend(data.get('value', []))

            # Update the endpoint for the next page, if any
            endpoint = data.get('@odata.nextLink')
            query_parameters = None  # Ensure subsequent requests do not repeat the initial query parameters

        if messages:
            sorted_messages = sorted(messages, key=lambda x: x['receivedDateTime'], reverse=True)

            if only_return_latest:
                return sorted_messages[0]['id']
            else:
                return sorted_messages

        raise Exception("No matching messages found.")
        
    def get_attachment_ids_and_names(self, message_id, shared_mailbox_email=None):
        """
        Retrieves a list of attachment IDs and names for a specified email message.

        This method uses the Microsoft Graph API to fetch all attachments for a given
        email message. It returns a list of tuples, each containing the ID and name of
        an attachment.

        Parameters:
            message_id (str): The ID of the email message whose attachments are to be retrieved.
            shared_mailbox_email (str): Optional. The email address of the shared mailbox to retrieve attachments from.

        Returns:
            list of tuple: A list of tuples, where each tuple contains the ID and name of an attachment.
        """
        # Determine the user ID or shared mailbox email to use in the endpoint
        mailbox_id = shared_mailbox_email if shared_mailbox_email else self.user_id

        # Constructing the URL for the Microsoft Graph API attachments endpoint
        url = f'https://graph.microsoft.com/v1.0/users/{mailbox_id}/messages/{message_id}/attachments'

        # Setting up the authorization header with the access token
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

        # Performing a GET request to retrieve the attachments
        response = requests.get(url, headers=headers)

        # Initializing an empty list to store attachment information
        attachments_info = []

        # Checking the response status code
        if response.status_code == 200:
            # Parsing the JSON response to extract attachment data
            attachments = response.json()
            
            # Extracting attachment IDs and names and adding them to the list
            attachments_info = [
                (attachment['id'], attachment['name'])
                for attachment in attachments.get('value', [])
            ]
        else:
            # Raising an exception if the API call was unsuccessful
            response.raise_for_status()

        return attachments_info
    
    def download_attachment(self, message_id, attachment_id, is_binary=False, encoding='utf-8', shared_mailbox_email=None):
        """
        Downloads an attachment from an email using the Microsoft Graph API.

        This method fetches an attachment from a specific email. It can handle both
        binary files (like images or PDFs) and text-based files (like CSV or TXT).

        Parameters:
            message_id (str): The ID of the message from which the attachment is to be downloaded.
            attachment_id (str): The ID of the attachment to be downloaded.
            is_binary (bool): Flag indicating whether the attachment is a binary file. Default is False.
            encoding (str): The encoding used to decode the attachment's content. Ignored if is_binary is True. Default is 'utf-8'.
            shared_mailbox_email (str): Optional. The email address of the shared mailbox to retrieve the attachment from.

        Returns:
            bytes or str: The content of the attachment, either as a byte string (for binary files) or a decoded string (for text files).
        """
        # Determine the user ID or shared mailbox email to use in the endpoint
        mailbox_id = shared_mailbox_email if shared_mailbox_email else self.user_id

        # Constructing the URL to access the attachment via Microsoft Graph API
        attachment_url = f'https://graph.microsoft.com/v1.0/users/{mailbox_id}/messages/{message_id}/attachments/{attachment_id}'

        # Setting up the authorization header with the access token
        headers = {'Authorization': f'Bearer {self.access_token}'}

        # Making a GET request to fetch the attachment and check success
        attachment_response = requests.get(attachment_url, headers=headers)
        attachment_response.raise_for_status()

        # Extracting the attachment content from the response
        attachment = attachment_response.json()
        attachment_content = base64.b64decode(attachment["contentBytes"])

        # Returning the attachment content based on its type
        return attachment_content if is_binary else attachment_content.decode(encoding)

    def delete_email(self, message_id, shared_mailbox_email=None):
        """
        Deletes an email using the Microsoft Graph API.

        This method sends a DELETE request to the Microsoft Graph API to remove an
        email message from a user's mailbox or a shared mailbox. It uses the user's ID
        or the shared mailbox email and the message's ID to identify the specific email to be deleted.

        Parameters:
            message_id (str): The ID of the message to be deleted.
            shared_mailbox_email (str): Optional. The email address of the shared mailbox to delete the email from.

        Returns:
            bool: True if the email was successfully deleted. Raises an exception otherwise.
        """
        # Determine the user ID or shared mailbox email to use in the endpoint
        mailbox_id = shared_mailbox_email if shared_mailbox_email else self.user_id

        # Constructing the URL for the Microsoft Graph API delete message endpoint
        delete_url = f'https://graph.microsoft.com/v1.0/users/{mailbox_id}/messages/{message_id}'

        # Setting up the authorization header with the access token
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

        # Sending the DELETE request
        response = requests.delete(delete_url, headers=headers)

        # Check if the request was successful
        if response.status_code == 204:
            print("Email deleted successfully.")
            return True
        else:
            raise Exception(f"Failed to delete email. Status code: {response.status_code}, Response: {response.text}")
    
    def get_user(self):
        """
        Gets the user id of the currently authenticated user using the Microsoft Graph API.

        Returns:
            json: user_data

        To access items in the json use user_data.get('id')
        """

        # Constructing the URL for the Microsoft Graph API me endpoint
        me_url = "https://graph.microsoft.com/v1.0/me"

        # Setting up the authorization header with the access token
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }

        # GET Response from endpoint.
        response = requests.get(me_url, headers=headers)
        response.raise_for_status()

    # Extract user data
        user_data = response.json()

        return user_data
    
    def send_email(self, subject, content, email_addresses):
        """
        Sends an email using the Microsoft Graph API.
        :param subject: Email subject.
        :param content: Email content.
        :param recipients: A list of recipient email addresses.
        """
        url = 'https://graph.microsoft.com/v1.0/me/sendMail'
        headers = {'Authorization': f'Bearer {self.access_token}', 'Content-Type': 'application/json'}

        # Produce recipients section in correct format.
        recipients = []
        for address in email_addresses:
            recipients.append({"emailAddress": {"address": address}})

        # JSON Representing the email.
        email = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": content
                },
                "toRecipients": recipients
            },
            "saveToSentItems": "true"
        }

        # Post the email.
        response = requests.post(url, headers=headers, data=json.dumps(email))

        if response.status_code == 202:
            print("Email sent successfully!")
        else:
            print(f"Failed to send email. Status code: {response.status_code}")
            raise Exception(f"Failed to send email. Status code: {response.status_code}")
    
    def _check_if_message_exists(self, message_id, shared_mailbox_email=None):
        """
        Checks if an email message exists in the user's mailbox or a shared mailbox.

        Parameters:
            message_id (str): The ID of the message to be checked.
            shared_mailbox_email (str): Optional. The email address of the shared mailbox to check the message in.

        Returns:
            bool: True if the message exists, False otherwise.
        """
        # Determine the user ID or shared mailbox email to use in the endpoint
        mailbox_id = shared_mailbox_email if shared_mailbox_email else self.user_id

        # Constructing the URL for the Microsoft Graph API message endpoint
        message_url = f'https://graph.microsoft.com/v1.0/users/{mailbox_id}/messages/{message_id}/'

        # Setting up the authorization header with the access token
        headers = {'Authorization': f'Bearer {self.access_token}'}

        try:
            # Making a GET request to the Graph API
            response = requests.get(message_url, headers=headers)

            # If the status code is 200, the message exists
            if response.status_code == 200:
                return True
            else:
                return False
        except Exception as e:
            # Log any exceptions
            raise Exception(f"Failed to check for message existence. Error: {e}") from None

