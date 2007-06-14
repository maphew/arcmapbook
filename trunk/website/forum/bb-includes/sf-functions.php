<?php
/* Additional functions for enabling mail on sourceforge
 * by Pravin Paratey 
 * http://pravin.insanitybegins.com/articles/running-bbpress-on-sourceforge/
 *
 *
 * Note: I use the same database as bbpress. I create an additional table
 * called sfMailTable
 * 
 */
 
 
// This function pushes data into the table
function sfmail ($email, $subject, $message) {
	global $bbdb;
	$table = 'sfMailTable';
 
	// Create table if it does not exist
	if($bbdb->get_var("SHOW TABLES LIKE '$table'") != $table) {
		$sql = "CREATE TABLE $table (
			id	bigint	not null auto_increment,
			email	text	not null,
			subject	text	not null,
			message	text	not null,
			unique key id(id)
			);";
		$results = $bbdb->query($sql);
	}
 
	// Push email data into the table
	$results = $bbdb->query("INSERT INTO `$table` (email, subject, message)" .
		"VALUES ('$email', '$subject', '$message');");
}
?>