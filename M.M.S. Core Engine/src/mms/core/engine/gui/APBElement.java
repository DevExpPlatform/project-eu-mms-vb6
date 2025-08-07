package mms.core.engine.gui;

import org.eclipse.swt.SWT;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.swt.widgets.Label;
import org.eclipse.swt.widgets.ProgressBar;
import org.eclipse.wb.swt.SWTResourceManager;

import mms.core.engine.utils.Utils;

public class APBElement extends Composite {

	private Label 		lblETCETA		= null;
	private Label 		lblInfoLeft		= null;
	private Label 		lblInfoRight	= null;
	private Label 		lblStatus		= null;
	private ProgressBar pbItem			= null;

	private int 		maxValue 		= 100;
	private int 		maxValueETCETA  = 100;
	private long 	  	startTime 		= 0; 
	private String 	  	strETA    		= "";
	private int 	  	tickCntr		= 0;
	private int 	  	tickness      	= 1000;
	private String 		tmpCaptionLeft 	= "";
	private String 		tmpCaptionRight = "";
	private int       	tmpJobStatus 	= -1;
	private int 		tmpValue 		= -1;

	public APBElement(Composite parent, int style) {
		super(parent, style);

		GridLayout gridLayout = new GridLayout(2, false);
		gridLayout.verticalSpacing = 3;
		gridLayout.marginWidth = 0;
		gridLayout.marginHeight = 0;
		gridLayout.horizontalSpacing = 3;
		setLayout(gridLayout);

		Label lblHStripe = new Label(this, SWT.SEPARATOR | SWT.HORIZONTAL);
		lblHStripe.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, false, false, 2, 1));
		
		lblETCETA = new Label(this, SWT.NONE);
		lblETCETA.setFont(SWTResourceManager.getFont("Tahoma", 8, SWT.BOLD));
		lblETCETA.setLayoutData(new GridData(SWT.LEFT, SWT.CENTER, true, false, 1, 1));
		lblETCETA.setText("ETC: Unknown - ETA: Unknown");
		
		GridData gd_lblStatus = new GridData(SWT.FILL, SWT.CENTER, false, false, 1, 1);
		gd_lblStatus.widthHint = 30;
		
		lblStatus = new Label(this, SWT.RIGHT);
		lblStatus.setFont(SWTResourceManager.getFont("Microsoft Sans Serif", 8, SWT.BOLD));
		lblStatus.setForeground(SWTResourceManager.getColor(SWT.COLOR_DARK_RED));
		lblStatus.setLayoutData(gd_lblStatus);
		lblStatus.setText("0%");

		pbItem = new ProgressBar(this, SWT.NONE);
		pbItem.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 2, 1));

		GridLayout gl_cmpInfoLabels = new GridLayout(2, false);
		gl_cmpInfoLabels.verticalSpacing = 3;
		gl_cmpInfoLabels.marginWidth = 3;
		gl_cmpInfoLabels.marginHeight = 3;
		gl_cmpInfoLabels.horizontalSpacing = 3;
	
		Composite cmpInfoLabels = new Composite(this, SWT.NONE);
		cmpInfoLabels.setLayout(gl_cmpInfoLabels);
		cmpInfoLabels.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 2, 1));

		lblInfoLeft = new Label(cmpInfoLabels, SWT.NONE);
		lblInfoLeft.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));

		lblInfoRight = new Label(cmpInfoLabels, SWT.NONE);
		lblInfoRight.setAlignment(SWT.RIGHT);
		lblInfoRight.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));

		this.pack();
	}

	@Override
	protected void checkSubclass() {}

	public void getETCETA(int value) {
		if (this.tmpValue == -1) {
			this.startTime 	  = System.currentTimeMillis();
			this.tickCntr  	  = 0;
			this.tmpJobStatus = -1;
			
			lblETCETA.setText("ETC: Unknown - ETA: Unknown"); 
			pbItem.setSelection(0);
		} 
	
		long elapsedTime = System.currentTimeMillis() - startTime;
		int  jobStatus   = (int) ((value * tickness) / maxValueETCETA);

		if (tmpJobStatus != jobStatus) {
			tmpJobStatus  =  jobStatus;
			tickCntr     += 1;
	
			strETA = Utils.ms2HMS((elapsedTime / tickCntr) * (tickness - tmpJobStatus));
		}

		lblETCETA.setText("ETC: " + Utils.ms2HMS(elapsedTime) + " - ETA: " + strETA);
	}

	public void setETCETAMaxValue(int value) {
		this.maxValueETCETA = value;
	}

	public void setLabelLeftCaption(String strCaption) {
		if (!tmpCaptionLeft.equals(strCaption)) {
			lblInfoLeft.setText(strCaption);
			
			tmpCaptionLeft = strCaption;
		}
	}

	public void setLabelRightCaption(final String strCaption) {
		if (!tmpCaptionRight.equals(strCaption)) {
			lblInfoRight.setText(strCaption);
			
			tmpCaptionRight = strCaption;
		}
	}

	public void setMaximum(int value) {
		this.maxValue       = value;
		this.maxValueETCETA = value;
		this.tmpValue       = -1;
	}

	public void setSelection(int value) {
		if ((value != 0) && (tmpValue != value)) {
			int myValue = (int) ((value * 100) / maxValue);

			lblStatus.setText(myValue + "%");
			pbItem.setSelection(myValue);
			
			tmpValue = value;
		}
	}

}
