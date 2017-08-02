package com.yo.main;

import java.util.Date;

public class SnInfo {
	private String isn;
	private String activeTime;
	public String getIsn() {
		return isn;
	}
	public void setIsn(String isn) {
		this.isn = isn;
	}
	public String getActiveTime() {
		return activeTime;
	}
	public void setActiveTime(String activeTime) {
		this.activeTime = activeTime;
	}
	public SnInfo() {
		super();
		// TODO Auto-generated constructor stub
	}
	public SnInfo(String isn, String activeTime) {
		super();
		this.isn = isn;
		this.activeTime = activeTime;
	}
	@Override
	public String toString() {
		return "SnInfo [isn=" + isn + ", activeTime=" + activeTime + "]";
	}
	
}
