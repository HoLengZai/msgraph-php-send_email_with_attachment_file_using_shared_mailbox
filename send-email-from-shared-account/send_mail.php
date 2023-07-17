<?php
require_once __DIR__ . '/vendor/autoload.php';
$dotenv = Dotenv\Dotenv::createImmutable(__DIR__ . '/../');
$dotenv->load();

# APP UPLOADER
# To be filled for connecting to MS Graph
$tenant_id = $_ENV['TENANT_ID'];
$client_id = $_ENV['CLIENT_ID'];
$client_secret = $_ENV['CLIENT_SECRET'];
$send_as_mailbox_address = $_ENV['SEND_AS_MAILBOX_ADDRESS'];

# Set alias
use GuzzleHttp\Client;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model;

# Init GuzzleHttp Client class
$guzzle = new Client();

# Prepare URI for MS Graph Authentication
$url = 'https://login.microsoftonline.com/' . $tenant_id . '/oauth2/v2.0/token';
$token = json_decode($guzzle->post($url, [
	'form_params' => [
		'client_id' => $client_id,
		'client_secret' => $client_secret,
		'scope' => 'https://graph.microsoft.com/.default',
		'grant_type' => 'client_credentials',
	],
])->getBody()->getContents());

# Save MS Graph Authentication Access Token
$accessToken = $token->access_token;

# Init MS Graph class
$graph = new Graph();

# Set MS Graph Access Token retrieved for your $graph instance 
$graph->setAccessToken($accessToken);
printf("MS Graph Access Token Set: OK\n");

# E-mail subject
$mail_subject = "Test subject";

# E-mail TO field (example with 2 recipients)
$mail_toRecipients = array(
	array('emailAddress' => array(
		'name' => 'Display Name E-APP',
		'address' => '')
	),
	array('emailAddress' => array(
		'name' => 'Display Name EXT',
		'address' => '')
	),
);

# E-mail CC field (example with 2 recipients)
$mail_ccRecipients = array(
	array('emailAddress' => array(
		'name' => 'Display Name E-APP',
		'address' => '')
	),
	array('emailAddress' => array(
		'name' => 'Display Name EXT',
		'address' => '')
	),
);
		// ),     // name is optional, otherwise array of address=>email@address

$mail_importance = 'normal';

$mail_body = array (
	'contentType' => "HTML", //Text or HTML
	'content' => "This is a <b>Test</b>."
);

# To add an attachment in your e-mail. If you want inline image, you can specify the "cid" in your HTML tag of your e-mail body
$contentBytes = base64_encode(file_get_contents('.\hello.jpg'));
$mail_attachments = array(
	array('@odata.type' => '#microsoft.graph.fileAttachment', 'Name' => 'hello.jpg', 'ContentType' => 'application/jpeg', 'ContentBytes' =>  $contentBytes, 'ContentID' => 'cid:hello')
);

$mailArgs =  array(
	'message' => array(
		'subject' => $mail_subject,
		'toRecipients' => $mail_toRecipients,
		'ccRecipients' => $mail_ccRecipients,
		'importance' => $mail_importance,
		'conversationId' => '',   //optional, use if replying to an existing email to keep them chained properly in outlook
		'body' => $mail_body,
		'attachments' => $mail_attachments
	),
	"saveToSentItems" => "true"
);

// For DEBUG only
#print_r($mailArgs);

$msgJSON = json_encode($mailArgs);

// For DEBUG only
#print_r($msgJSON);

# Send the e-mail
$res = $graph->createRequest("POST", "https://graph.microsoft.com/v1.0/users/$send_as_mailbox_address/sendMail")
	->attachBody($msgJSON)
	->execute();

// For DEBUG only
printf("Send e-mail done\n");
