		</div>

		<!--
			If you like showing off the fact that your server rocks,
			<h3><?php bb_timer_stop(1); ?> - <?php echo $bbdb->num_queries; ?> queries</h3>
		-->

	</div>

	<div id="footer">
		<p><?php printf(__('%1$s powered by <a href="%2$s">bbPress</a>.</p>'), bb_option('name'), "http://bbpress.org") ?>
	</div>

	<?php do_action('bb_foot', ''); ?>
<a href="http://sourceforge.net"><img src="http://sflogo.sourceforge.net/sflogo.php?group_id=170350&amp;type=1" width="88" height="31" border="0" alt="SourceForge.net Logo" align="left" /></a> Web hosting and source code management services donated by Sourceforge.net

</body>
</html>
