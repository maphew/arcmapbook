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

</body>
</html>
