package com.soul;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.util.*;
class OperatorView extends JFrame {

    private JButton generateButton;

    private JPanel controlPanel = new JPanel();

    private File[] fileArray;

    private JLabel fileNumLabel;

    OperatorView(String title) {
        super(title);

        addSubviews();
    }

    private void addSubviews() {
        controlPanel.setLayout(null);
        this.add(controlPanel);

//        fileNumLabel = new JLabel("您选择了" + fileArray.length + "个文件");
        fileNumLabel = new JLabel("您还没有选择任何文件！");
        fileNumLabel.setBounds(0, 340, 300, 40);
        fileNumLabel.setHorizontalAlignment(SwingConstants.CENTER);
        controlPanel.add(fileNumLabel);

        JButton button = new JButton("选择文件");
        button.setBounds(10, 300, 100, 40);
        controlPanel.add(button);
        button.setActionCommand("selectFiles");
        button.addActionListener(new ButtonClickListener());
    }

    private class ButtonClickListener implements ActionListener {

        public void actionPerformed(ActionEvent e) {
            String command = e.getActionCommand();
            if (command.equals("selectFiles")) {
                System.out.print("选择文件！！！");
                selectFiles();
            }else if (command.equals("generateFile")) {
                System.out.print("生成文件！！！");
            }
        }
    }

    private void selectFiles() {
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "word", "doc", "docx");
        fileChooser.setFileFilter(filter);
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setMultiSelectionEnabled(true);
        try {
            int returnVal = fileChooser.showOpenDialog(this);
            if(returnVal == JFileChooser.APPROVE_OPTION) {
                fileArray = fileChooser.getSelectedFiles();
                layoutViews();
                System.out.println("You choose "+ fileChooser.getSelectedFiles().length + "  files" );
                ArrayList<String> paths = new ArrayList<String>();
                for (int i = 0; i < fileChooser.getSelectedFiles().length; i++) {
                    paths.add(fileChooser.getSelectedFiles()[i].getAbsolutePath());
                }
                WordRead reader = new WordRead(paths);

            }
        }catch (HeadlessException exception) {
            exception.printStackTrace();
        }

    }

    private void layoutViews() {

        fileNumLabel.setText("您选择了" + fileArray.length + "个文件");

        generateButton = new JButton("生成文件");
        generateButton.setBounds(180, 300, 100, 40);
        controlPanel.add(generateButton);
        generateButton.setActionCommand("generateFile");
        generateButton.addActionListener(new ButtonClickListener());

    }
}
