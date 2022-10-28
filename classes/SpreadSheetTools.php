<?php

namespace Classes;

use Exception;
use Google\Client;
use Google\Service\Drive;
use Google\Service\Sheets;
use Google\Service\Sheets\Spreadsheet;
use Google\Service\Sheets\ValueRange;

class SpreadSheetTools
{
    private $client;

    public function __construct()
    {
        $this->client = $this->getClient();
    }
   
    private function getClient()
    {
        $client = new Client();
        $client->setScopes(Drive::DRIVE);
        $client->setAuthConfig('credentials.json');
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        $tokenPath = 'token.json';
        if (file_exists($tokenPath)) {
            $accessToken = json_decode(file_get_contents($tokenPath), true);
            $client->setAccessToken($accessToken);
        }

        if ($client->isAccessTokenExpired()) {
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            } else {
                $authUrl = $client->createAuthUrl();
                printf("Open the following link in your browser:\n%s\n", $authUrl);
                print 'Enter verification code: ';
                $authCode = trim(fgets(STDIN));

                $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
                $client->setAccessToken($accessToken);

                if (array_key_exists('error', $accessToken)) {
                    throw new Exception(join(', ', $accessToken));
                }
            }
         
            if (!file_exists(dirname($tokenPath))) {
                mkdir(dirname($tokenPath), 0700, true);
            }
            file_put_contents($tokenPath, json_encode($client->getAccessToken()));
        }
        return $client;
    }

    
    public function create($title)
    {   
        $service = new Sheets($this->client);
        try{
            $spreadsheet = new Spreadsheet([
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
            echo 'Message: '.$e->getMessage();
        }
    }

    public function updateValues($spreadsheetId, $range, $valueInputOption, $values)
    {
        $service = new Sheets($this->client);
        try{
            $body = new ValueRange([
                'values' => $values
            ]);
            $params = [
                'valueInputOption' => $valueInputOption
            ];
            $result = $service->spreadsheets_values->update($spreadsheetId, $range, $body, $params);
            printf("%d cells updated.", $result->getUpdatedCells());
            return $result;
        }
        catch(Exception $e) {
            echo 'Message: '.$e->getMessage();
        }
    }
}