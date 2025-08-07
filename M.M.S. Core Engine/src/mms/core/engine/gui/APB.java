package mms.core.engine.gui;

import java.util.ArrayList;
import java.util.List;

import org.eclipse.swt.SWT;
import org.eclipse.swt.graphics.Rectangle;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Group;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.wb.swt.SWTResourceManager;

public class APB {

	private  String 	  		frmCaption		= "A.P.B.";
	private  int 				numProgressBar 	= 1;

	protected String 			errMsg			= "";
	protected Shell 	  		frmMainShell	= null;
	protected boolean 			jobRunning		= false;
	protected List<APBElement> 	myAPBList 		= new ArrayList<APBElement>();
	protected boolean 			rValue 			= false;
	protected boolean 			unattendedMode	= false;

	/**
	 * @wbp.parser.entryPoint
	 */
	private void createMainContents() {
		GridLayout gl_frmMainShell = new GridLayout(1, false);
		gl_frmMainShell.verticalSpacing = 3;
		gl_frmMainShell.marginWidth = 3;
		gl_frmMainShell.marginHeight = 3;
		gl_frmMainShell.horizontalSpacing = 3;
	
		frmMainShell = new Shell(SWT.TITLE | SWT.SYSTEM_MODAL | SWT.ON_TOP);
		frmMainShell.setText(frmCaption);
		frmMainShell.setLayout(gl_frmMainShell);

		GridLayout gl_shell = new GridLayout(1, false);
		gl_shell.verticalSpacing = 3;
		gl_shell.marginWidth = 3;
		gl_shell.marginHeight = 3;
		gl_shell.horizontalSpacing = 3;
	
		GridLayout gl_grpWorkInProgress = new GridLayout(1, false);
		gl_grpWorkInProgress.verticalSpacing = 2;
		gl_grpWorkInProgress.marginWidth = 2;
		gl_grpWorkInProgress.marginHeight = 2;
		gl_grpWorkInProgress.horizontalSpacing = 2;
	
		GridData gd_grpWorkInProgress = new GridData(SWT.FILL, SWT.FILL, true, true, 1, 1);
		gd_grpWorkInProgress.widthHint = 380;
	
		Group grpWorkInProgress = new Group(frmMainShell, SWT.NONE);
		grpWorkInProgress.setFont(SWTResourceManager.getFont("Microsoft Sans Serif", 8, SWT.BOLD));
		grpWorkInProgress.setText("Work in Progress:");
		grpWorkInProgress.setLayout(gl_grpWorkInProgress);
		grpWorkInProgress.setLayoutData(gd_grpWorkInProgress);

		for (int i = 0; i < numProgressBar; i++) {
			APBElement myAPB = new APBElement(grpWorkInProgress, SWT.NONE);
			myAPB.setLayoutData(new GridData(SWT.FILL, SWT.CENTER, true, false, 1, 1));
			
			myAPBList.add(myAPB);
		}
		
		/*
		 * Form Positioning 
		 */
		Rectangle Screen = Display.getCurrent().getBounds();

		frmMainShell.pack();
		frmMainShell.setLocation((Screen.width - frmMainShell.getBounds().width) / 2, (Screen.height - frmMainShell.getBounds().height) / 2);
	}

	public String getErrMsg() {
		return errMsg;
	}

	/**
	 * @wbp.parser.entryPoint
	 */
	public boolean open() {
		if (unattendedMode) {
			jobExecutor();
		} else {
			Display display = Display.getDefault();

			createMainContents();

			frmMainShell.open();
			frmMainShell.layout();

			jobRunning = true;

			jobExecutor();
			
			while (!frmMainShell.isDisposed()) {
				if (!display.readAndDispatch()) {
					display.sleep();
				}
			}
		}

		return rValue;
	}

	protected void jobExecutor() {}

	public void setCaption(String frmCaption) {
		this.frmCaption = frmCaption;
		
		if (this.frmMainShell != null) 
			this.frmMainShell.setText(frmCaption);
	}

	public void setProgressNum(int numProgressBar) {
		this.numProgressBar = numProgressBar;
	}

	public void setUnattendedMode(boolean unattendedMode) {
		this.unattendedMode = unattendedMode;
	}

}
