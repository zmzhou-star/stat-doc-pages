package com.github.zmzhoustar;

import org.junit.Test;

import static org.junit.Assert.assertNotNull;

/**
 * Unit test for simple App.
 */
public class AppTest {
	/**
	 * Rigorous Test :-)
	 */
	@Test
	public void statFileNum() {
		Long total = App.statFileNum("C:\\tmp\\demo");
		assertNotNull(total);
	}
}
