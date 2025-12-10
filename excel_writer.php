<?php

header('Content-Type: application/json');
require_once 'vendor/autoload.php';

$dotenv = Dotenv\Dotenv::createImmutable(__DIR__);
$dotenv->load();

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Method not allowed']);
    exit;
}

$input = json_decode(file_get_contents('php://input'), true);

if (!isset($input['secret']) || $input['secret'] !== $_ENV['API_SECRET']) {
    http_response_code(401);
    echo json_encode(['error' => 'Invalid secret token']);
    exit;
}

if (!isset($input['FirstName']) || !isset($input['LastName']) || !isset($input['Email'])) {
    http_response_code(400);
    echo json_encode(['error' => 'Missing required fields: FirstName, LastName, Email']);
    exit;
}

$firstName = $input['FirstName'];
$lastName = $input['LastName'];
$email = $input['Email'];

try {
    // Get OAuth token
    $tokenUrl = "https://login.microsoftonline.com/{$_ENV['TENANT_ID']}/oauth2/v2.0/token";

    $tokenData = [
        'client_id' => $_ENV['CLIENT_ID'],
        'client_secret' => $_ENV['CLIENT_SECRET'],
        'scope' => 'https://graph.microsoft.com/.default',
        'grant_type' => 'client_credentials'
    ];

    $ch = curl_init($tokenUrl);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($tokenData));

    $tokenResponse = curl_exec($ch);
    curl_close($ch);

    $tokenJson = json_decode($tokenResponse, true);

    if (!isset($tokenJson['access_token'])) {
        throw new Exception('Failed to get access token: ' . json_encode($tokenJson));
    }

    $accessToken = $tokenJson['access_token'];

    $siteId = $_ENV['SITE_ID'];
    $driveId = $_ENV['DRIVE_ID'];
    $sourceFileId = $_ENV['SOURCE_FILE_ID'];
    $worksheetName = $_ENV['WORKSHEET_NAME'];

    // Copy the source file
    $timestamp = date('YmdHis');
    $newFileName = "UserData_{$firstName}_{$lastName}_{$timestamp}.xlsx";

    $copyUrl = "https://graph.microsoft.com/v1.0/sites/{$siteId}/drives/{$driveId}/items/{$sourceFileId}/copy";

    $copyData = json_encode([
        'name' => $newFileName,
        'parentReference' => [
            'driveId' => $driveId
        ]
    ]);

    $ch = curl_init($copyUrl);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_POST, true);
    curl_setopt($ch, CURLOPT_POSTFIELDS, $copyData);
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        "Authorization: Bearer {$accessToken}",
        'Content-Type: application/json'
    ]);

    $copyResponse = curl_exec($ch);
    $copyHttpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($copyHttpCode !== 202) {
        throw new Exception('Failed to copy file. HTTP Code: ' . $copyHttpCode . ', Response: ' . $copyResponse);
    }

    // Wait for copy operation to complete and retry search
    $maxRetries = 10;
    $retryCount = 0;
    $newFileId = null;

    while ($retryCount < $maxRetries) {
        sleep(2);

        $searchUrl = "https://graph.microsoft.com/v1.0/sites/{$siteId}/drives/{$driveId}/root/search(q='{$newFileName}')";

        $ch = curl_init($searchUrl);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
        curl_setopt($ch, CURLOPT_HTTPHEADER, [
            "Authorization: Bearer {$accessToken}"
        ]);

        $searchResponse = curl_exec($ch);
        curl_close($ch);

        $searchJson = json_decode($searchResponse, true);

        if (isset($searchJson['value'][0]['id'])) {
            $newFileId = $searchJson['value'][0]['id'];
            break;
        }

        $retryCount++;
    }

    if (!$newFileId) {
        throw new Exception('Could not find copied file after ' . $maxRetries . ' attempts');
    }

    // Update the Excel file with user data
    $updateUrl = "https://graph.microsoft.com/v1.0/sites/{$siteId}/drives/{$driveId}/items/{$newFileId}/workbook/worksheets/{$worksheetName}/range(address='A1:B3')";

    $updateData = json_encode([
        'values' => [
            ['FirstName', $firstName],
            ['LastName', $lastName],
            ['Email', $email]
        ]
    ]);

    $ch = curl_init($updateUrl);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'PATCH');
    curl_setopt($ch, CURLOPT_POSTFIELDS, $updateData);
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        "Authorization: Bearer {$accessToken}",
        'Content-Type: application/json'
    ]);

    $updateResponse = curl_exec($ch);
    $updateHttpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($updateHttpCode !== 200) {
        throw new Exception('Failed to update Excel file. HTTP Code: ' . $updateHttpCode . ', Response: ' . $updateResponse);
    }

    echo json_encode([
        'success' => true,
        'message' => 'Excel file created and updated successfully',
        'fileName' => $newFileName,
        'fileId' => $newFileId
    ]);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode([
        'error' => 'Server error: ' . $e->getMessage()
    ]);
}
