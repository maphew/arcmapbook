<?php
// 2007 June 14, matt wilkie
// adapted for Sourceforge as per  http://pravin.insanitybegins.com/articles/running-bbpress-on-sourceforge/
include_once('sf-functions.php');

function bb_verify_email( $email ) {
	if (ereg('^[-!#$%&\'*+\\./0-9=?A-Z^_`a-z{|}~]+'.'@'.
		'[-!#$%&\'*+\\/0-9=?A-Z^_`a-z{|}~]+\.'.
		'[-!#$%&\'*+\\./0-9=?A-Z^_`a-z{|}~]+$', $email)) {
		if ( $check_domain && function_exists('checkdnsrr') ) {
			list (, $domain)  = explode('@', $email);
			if ( checkdnsrr($domain . '.', 'MX') || checkdnsrr($domain . '.', 'A') ) {
				$r = $email;
			} else {
				$r = false;
			}
		} else {
			$r = $email;
		}
	} else {
		$r = false;
	}
	return apply_filters( 'bb_verify_email', $r, $email );
}

function bb_update_user( $user_id, $email, $url ) {
	global $bbdb, $bb_cache;

	$bbdb->query("UPDATE $bbdb->users SET
	user_email = '$email',
	user_url   = '$url'
	WHERE ID   = '$user_id'
	");
	$bb_cache->flush_one( 'user', $user_id );

	do_action('bb_update_user', $user_id);
	return $user_id;
}

function bb_reset_email( $user_login ) {
	global $bbdb;
	$user = $bbdb->get_row("SELECT * FROM $bbdb->users WHERE user_login = '$user_login'");

	$resetkey = bb_random_pass( 15 );
	bb_update_usermeta( $user->ID, 'newpwdkey', $resetkey );
	if ( $user ) :
		sfmail( bb_get_user_email( $user->ID ), bb_get_option('name') . ': ' . __('Password Reset'), sprintf( __("If you wanted to reset your password, you may do so by visiting the following address:

%s

If you do not want to reset your password, just ignore this email. Thanks!"), bb_get_option('uri')."bb-reset-password.php?key=".$resetkey ), 'From: ' . bb_get_option('admin_email') );

	endif;
}

function bb_reset_password( $key ) {
	global $bbdb;
	$key = user_sanitize( $key );
	if ( empty( $key ) )
		bb_die(__('Key not found.'));
	if ( !$user_id = $bbdb->get_var("SELECT user_id FROM $bbdb->usermeta WHERE meta_key = 'newpwdkey' AND meta_value = '$key'") )
		bb_die(__('Key not found.'));
	if ( $user = new BB_User( $user_id ) ) :
		if ( bb_has_broken_pass( $user->ID ) )
			bb_block_current_user();
		if ( !$user->has_cap( 'change_password' ) )
			bb_die( __('You are not allowed to change your password.') );
		$newpass = bb_random_pass( 6 );
		bb_update_user_password( $user->ID, $newpass );
		bb_send_pass           ( $user->ID, $newpass );
		bb_update_usermeta( $user->ID, 'newpwdkey', '' );
	else :
		bb_die(__('Key not found.'));
	endif;
}

function bb_update_user_password( $user_id, $password ) {
	global $bbdb, $bb_cache;
	$passhash = md5( $password );

	$bbdb->query("UPDATE $bbdb->users SET
	user_pass = '$passhash'
	WHERE ID = '$user_id'
	");
	$bb_cache->flush_one( 'user', $user_id );

	do_action('bb_update_user_password', $user_id);
	return $user_id;
}

function bb_random_pass( $length = 6) {
	$number = mt_rand(1, 15);
	$string = md5( uniqid( microtime() ) );
 	$password = substr( $string, $number, $length );
	return $password;
}

function bb_send_pass( $user, $pass ) {
	global $bbdb;
	$user = (int) $user;
	$user = $bbdb->get_row("SELECT * FROM $bbdb->users WHERE ID = $user");

	if ( $user ) :
		$message = __("Your username is: %1\$s \nYour password is: %2\$s \nYou can now log in: %3\$s \n\nEnjoy!");
		sfmail( bb_get_user_email( $user->ID ), bb_get_option('name') . ': ' . __('Password'), 
			sprintf( $message, "$user->user_login", "$pass", bb_get_option('uri') ), 
			'From: ' . bb_get_option('admin_email') 
		);

	endif;
}
?>
