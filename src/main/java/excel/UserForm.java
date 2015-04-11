package excel;
import javax.swing.*;

import java.awt.*;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;

import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
public class UserForm  extends JPanel implements ActionListener{
	   
	 
	    public UserForm() {
	        super(new GridLayout(1,2));
	 
	        JButton button = new JButton("Test button");
	        JButton button2 = new JButton("Test 2");
	        button.addActionListener(this);
	        button2.addActionListener(this);
	        add(button);
	        add(button2);
	    }
	 
	  
	 
	    /**
	     * Create the GUI and show it.  For thread safety,
	     * this method should be invoked from the
	     * event-dispatching thread.
	     */
	    
	    public static int DEFAULT_WIDTH = 300;
	    public static int DEFAULT_HEIGHT = 300;
	   
	    private static void createAndShowGUI() {
	        //Create and set up the window.
	        JFrame frame = new JFrame("SimpleTableDemo");
	        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	        frame.setSize(DEFAULT_WIDTH, DEFAULT_HEIGHT);
	 
	        //Create and set up the content pane.
	        UserForm newContentPane = new UserForm();
	        newContentPane.setOpaque(true); //content panes must be opaque
	        frame.setContentPane(newContentPane);
	 
	        //Display the window.
	        frame.pack();
	        
	        frame.setVisible(true);
	    }
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		javax.swing.SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
	}



	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		System.out.println("Pressed");
	}

}
