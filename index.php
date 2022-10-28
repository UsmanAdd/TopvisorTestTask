<?php

require __DIR__ . '/vendor/autoload.php';

if (php_sapi_name() != 'cli') {
    throw new Exception('This application must be run on the command line.');
}

use Google\Client;
use Google\Service\Sheets\SpreadSheet;

/**
 * Returns an authorized API client.
 * @return Client the authorized client object.
 */
function getClient()
{
    $client = new Google\Client();
    $client->setApplicationName('Google Sheets API PHP Quickstart');
    $client->setScopes('https://www.googleapis.com/auth/spreadsheets');
    $client->setAuthConfig('credentials.json');
    $client->setAccessType('offline');
    $client->setPrompt('select_account consent');

    // Load previously authorized token from a file, if it exists.
    // The file token.json stores the user's access and refresh tokens, and is
    // created automatically when the authorization flow completes for the first
    // time.
    $tokenPath = 'token.json';
    if (file_exists($tokenPath)) {
        $accessToken = json_decode(file_get_contents($tokenPath), true);
        $client->setAccessToken($accessToken);
    }

    // If there is no previous token or it's expired.
    if ($client->isAccessTokenExpired()) {
        // Refresh the token if possible, else fetch a new one.
        if ($client->getRefreshToken()) {
            $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
        } else {
            // Request authorization from the user.
            $authUrl = $client->createAuthUrl();
            printf("Open the following link in your browser:\n%s\n", $authUrl);
            print 'Enter verification code: ';
            $authCode = trim(fgets(STDIN));

            // Exchange authorization code for an access token.
            $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
            $client->setAccessToken($accessToken);

            // Check to see if there was an error.
            if (array_key_exists('error', $accessToken)) {
                throw new Exception(join(', ', $accessToken));
            }
        }
        // Save the token to a file.
        if (!file_exists(dirname($tokenPath))) {
            mkdir(dirname($tokenPath), 0700, true);
        }
        file_put_contents($tokenPath, json_encode($client->getAccessToken()));
    }
    return $client;
}

/**
 * Create an empty spreadsheet.
 * 
 * Returns the created spreadsheet id.
 * @return Spreadsheet ID.
 */
 function create($title)
{   
    $client = getClient();
    $service = new Google_Service_Sheets($client);
    try{

        $spreadsheet = new Google_Service_Sheets_Spreadsheet([
            'properties' => [
                'title' => $title
                ]
            ]);
            $spreadsheet = $service->spreadsheets->create($spreadsheet, [
                'fields' => 'spreadsheetId'
            ]);
            printf("Spreadsheet ID: %s\n", $spreadsheet->spreadsheetId);
            return $spreadsheet->spreadsheetId;
    }
    catch(Exception $e) {
        // TODO(developer) - handle error appropriately
        echo 'Message: ' .$e->getMessage();
      }
}

function updateValues($spreadsheetId, $range, $valueInputOption, $values)
{
    $client = getClient();
    $service = new Google_Service_Sheets($client);
    try{
        $body = new Google_Service_Sheets_ValueRange([
        'values' => $values
        ]);
        $params = [
        'valueInputOption' => $valueInputOption
        ];
        //executing the request
        $result = $service->spreadsheets_values->update($spreadsheetId, $range, $body, $params);
        printf("%d cells updated.", $result->getUpdatedCells());
        return $result;
    }
    catch(Exception $e) {
        // TODO(developer) - handle error appropriately
        echo 'Message: ' .$e->getMessage();
    }
}

/**
 * Create a blank google sheet spreadsheet with a particular title.
 */
$spreadsheetId = create('My Test Google Sheet SpreadSheet');
updateValues($spreadsheetId,'A1:A10','RAW',[[1],[2],[3],[4],[5],[6],[7],[8],[9],[10]]);