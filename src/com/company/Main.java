package com.company;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.*;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.plaf.metal.DefaultMetalTheme;
import javax.swing.plaf.metal.MetalLookAndFeel;
import javax.swing.plaf.metal.OceanTheme;
import javax.swing.plaf.nimbus.NimbusLookAndFeel;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main extends JFrame {

    private static Main prog;

    private static NumberFormat nf = NumberFormat.getInstance();
    private static Map<String, Integer> map = new HashMap<>();       // карта для перевода символьного названия колонки в числовой

    private List<JLabel> example = new ArrayList<>();
    private List<JPanel> examplePanels = new ArrayList<>();

    private List<JRadioButton> type1 = new ArrayList<>();
    private List<JRadioButton> type2 = new ArrayList<>();

    private List<JRadioButton> type3 = new ArrayList<>();
    private List<JRadioButton> type4 = new ArrayList<>();

    private static List<JTextField> list1 = new ArrayList<>();
    private static List<JTextField> list2 = new ArrayList<>();
    private static List<JTextField> list3 = new ArrayList<>();

    private int count = 1;
    private String name;
    private String path;
    private String directory = "";

    private static String[] text = new String[6];

    private JPanel panelDate = new JPanel();

    private Object[][] tableData = {{"", "", "", "", "", "", ""}};               //пустая массив для начального вывода таблиц
    private Object[] headers = {"П/п", "Дата", "ИНН/Счет", "Наименование/ФИО", "Приход", "Расход", "Назначение"};  //заголовок для вывода таблиц

    private HSSFWorkbook hssfWorkbook;                               // рабочая книга Excel для старых версий
    private XSSFWorkbook xssfWorkbook;                               // рабочая книга Excel для новых версий

    private static JTextField day = new JTextField();                                          //день подачи заявления о банкротстве
    private static JTextField month = new JTextField();                                        //месяц подачи заявления о банкротстве
    private static JTextField year = new JTextField();                                         //год подачи заявления о банкротстве
    private static JTextField sum = new JTextField();                                          //балансовая стоимость
    private static JTextField percent = new JTextField();                                      //процент суммы по ИНН в балансовой стоимости

    private static JTextField dayStart = new JTextField();                                          //день подачи заявления о банкротстве
    private static JTextField monthStart = new JTextField();                                        //месяц подачи заявления о банкротстве
    private static JTextField yearStart = new JTextField();                                         //год подачи заявления о банкротстве
    private static JTextField dayEnd = new JTextField();                                          //день подачи заявления о банкротстве
    private static JTextField monthEnd = new JTextField();                                        //месяц подачи заявления о банкротстве
    private static JTextField yearEnd = new JTextField();                                         //год подачи заявления о банкротстве

    private JLabel fileName = new JLabel();                          //путь к файлу
    private JLabel info = new JLabel();                              //информация о файле (количество страниц)
    private JLabel labelDate = new JLabel("");                  //информация о корректности ввода даты подачи заявления о банкротстве
    private JLabel label1 = new JLabel("");                     //информация о количестве найденных сделок за месяц (или невозможности распознать дату операции из файла)
    private JLabel label2 = new JLabel("");                     //информация о количестве найденных сделок за полгода для второй вкладки (или невозможности распознать дату операции из файла)
    private JLabel label3 = new JLabel("");                     //информация о количестве найденных сделок за год (или невозможности распознать дату операции из файла)
    private JLabel label4 = new JLabel("");                     //информация о количестве найденных сделок за 3 года (или невозможности распознать дату операции из файла)
    private JLabel label5 = new JLabel("");                     //информация о количестве найденных сделок за 10 лет (или невозможности распознать дату операции из файла)
    private JLabel label6 = new JLabel("");                     //информация о количестве найденных сделок универсальный отбор (или невозможности распознать дату операции из файла)
    private JLabel label2Info = new JLabel("");                 //есть ли ошибки в анализе сделок по балансовой стоимости (за полгода)
    private JLabel labelOpenExcel = new JLabel("");                  //информация о корректности ввода даты подачи заявления о банкротстве
    private JLabel labelSaveData = new JLabel("");                  //информация о корректности ввода даты подачи заявления о банкротстве

    private JLabel labelSave1 = new JLabel("");                     //информация о количестве найденных сделок за месяц (или невозможности распознать дату операции из файла)
    private JLabel labelSave2 = new JLabel("");                     //информация о количестве найденных сделок за полгода для второй вкладки (или невозможности распознать дату операции из файла)
    private JLabel labelSave3 = new JLabel("");                     //информация о количестве найденных сделок за год (или невозможности распознать дату операции из файла)
    private JLabel labelSave4 = new JLabel("");                     //информация о количестве найденных сделок за 3 года (или невозможности распознать дату операции из файла)
    private JLabel labelSave5 = new JLabel("");                     //информация о количестве найденных сделок за 10 лет (или невозможности распознать дату операции из файла)
    private JLabel labelSave6 = new JLabel("");                     //информация о количестве найденных сделок универсальный отбор (или невозможности распознать дату операции из файла)


    private JPanel jpanel = new JPanel();                            //панель для первой вкладки основной формы
    private JPanel jpanelRez1 = new JPanel();                        //панель для второй вкладки основной формы
    private JPanel jpanelRez2 = new JPanel();                        //панель для третьей вкладки основной формы
    private JPanel jpanelRez3 = new JPanel();                        //панель для четвертой вкладки основной формы
    private JPanel jpanelRez4 = new JPanel();                        //панель для пятой вкладки основной формы
    private JPanel jpanelRez5 = new JPanel();                        //панель для шестой вкладки основной формы
    private JPanel jpanelRez6 = new JPanel();                        //панель для седьмой вкладки основной формы
    private JPanel jpanelRez = new JPanel();                         //панель для всех вкладок с результатами

    private JPanel panelSearch1 = new JPanel();                      //панель для вывода сделок за месяц
    private JPanel panelSearch2 = new JPanel();                      //панель для вывода сделок за полгода (вторая вкладка)
    private JPanel panelSearch3 = new JPanel();                      //панель для вывода сделок за год
    private JPanel panelSearch4 = new JPanel();                      //панель для вывода сделок за 3 года
    private JPanel panelSearch5 = new JPanel();                      //панель для вывода сделок за 10 лет
    private JPanel panelSearch6 = new JPanel();                      //панель для вывода сделок универсальный отбор

    private JTabbedPane tabbedPane = new JTabbedPane();              //вкладки для страниц
    private JTabbedPane tabbedPaneFirst = new JTabbedPane();         //вкладки для главной формы (получение данныч/результат)
    private JTabbedPane tabbedPaneResult = new JTabbedPane();        //вкладки для 3 видов сделок
    private JTabbedPane tabbedPane1 = new JTabbedPane();             //вкладки для сделок с предпочтением формы
    private JTabbedPane tabbedPane2 = new JTabbedPane();             //вкладки для сделок с неравн встречным исполнением

    private JButton open = new JButton("Выбрать файл");          //кнопка выбора файла
    private JButton openExcel = new JButton("Открыть файл");      //кнопка чтения файла
    private JButton saveData = new JButton("Сохранить полный анализ в Excel");      //кнопка чтения файла
    private JButton search1 = new JButton("Начать анализ");   //кнопка анализа сделок за месяц
    private JButton search2 = new JButton("Начать анализ");   //кнопка анализа сделок за полгода
    private JButton search3 = new JButton("Начать анализ");   //кнопка анализа сделок за год
    private JButton search4 = new JButton("Начать анализ");   //кнопка анализа сделок за 3 года
    private JButton search5 = new JButton("Начать анализ");   //кнопка анализа сделок за 10 лет
    private JButton search6 = new JButton("Начать анализ");   //кнопка анализа сделок за 10 лет
    private JButton saveData1 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за месяц
    private JButton saveData2 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за полгода
    private JButton saveData3 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за год
    private JButton saveData4 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за 3 года
    private JButton saveData5 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за 10 лет
    private JButton saveData6 = new JButton("Сохранить текущий результат анализа в Excel");   //кнопка анализа сделок за 10 лет

    private JScrollPane scrollPane1;                                 //прокрутка для первой таблицы с результатом (вторая вкладка)
    private JScrollPane scrollPane2;                                 //прокрутка для второй таблицы с результатом (третья вкладка)
    private JScrollPane scrollPane3;                                 //прокрутка для третьей таблицы с результатом (четвертая вкладка)
    private JScrollPane scrollPane4;                                 //прокрутка для четвертой таблицы с результатом (пятая вкладка)
    private JScrollPane scrollPane5;                                 //прокрутка для пятой таблицы с результатом (шестая вкладка)
    private JScrollPane scrollPane6;                                 //прокрутка для шестой таблицы с результатом (седьмая вкладка)

    private JTextArea areaINN = new JTextArea();                     //список ИНН для отбора
    private JTextArea areaPurpose = new JTextArea();                 //список ключевых слов в назначении операции для отбора
    private JTextArea areaFIO = new JTextArea();                 //список ключевых слов в назначении операции для отбора

    private List<JTextField> jrowStarts = new ArrayList<>();         //для каждой страницы файла поле ввода номера строки, с которой начинаются сделки
    private List<JTextField> jnums = new ArrayList<>();              //для каждой страницы файла поле ввода название колонки с порядковым номером
    private List<JTextField> jdateNums = new ArrayList<>();          //для каждой страницы файла поле ввода название колонки с датой сделки
    private List<JTextField> jINNNumsPlus = new ArrayList<>();       //для каждой страницы файла поле ввода название колонки с ИНН получателя
    private List<JTextField> jINNNumsMinus = new ArrayList<>();      //для каждой страницы файла поле ввода название колонки с ИНН плательщика
    private List<JTextField> jpurposeNums = new ArrayList<>();       //для каждой страницы файла поле ввода название колонки с назначением платежа
    private List<JTextField> jsumNumPlus = new ArrayList<>();        //для каждой страницы файла поле ввода название колонки с суммой платежа (приход)
    private List<JTextField> jsumNumMinus = new ArrayList<>();       //для каждой страницы файла поле ввода название колонки с суммой платежа (расход)
    private JButton[] helps = new JButton[8];                        //для каждого поля ввода кнопочка с примером ввода

    private List<JLabel> INNNumsPlus = new ArrayList<>();       //для каждой страницы файла поле ввода название колонки с ИНН получателя
    private List<JLabel> INNNumsMinus = new ArrayList<>();      //для каждой страницы файла поле ввода название колонки с ИНН плательщика

    private List<JLabel> sumNumsPlus = new ArrayList<>();       //для каждой страницы файла поле ввода название колонки с ИНН получателя
    private List<JLabel> sumNumsMinus = new ArrayList<>();      //для каждой страницы файла поле ввода название колонки с ИНН плательщика

    private List<JPanel> panelsHide = new ArrayList<>();            //для каждой страницы файла панелька с данными, которые указывает пользователь
    private List<JPanel> panelsDown = new ArrayList<>();            //для каждой страницы файла панелька правая (с загрузкой данных)
    private List<JPanel> panelsRight = new ArrayList<>();            //для каждой страницы файла панелька правая (с загрузкой данных)

    private List<JPanel> panelsRightTable = new ArrayList<>();      //для каждой страницы файла панелька с табличкой
    private List<JProgressBar> progressBars = new ArrayList<>();    //для каждой страницы файла прогресс бар загрузки данных
    private List<JScrollPane> scrollPanes = new ArrayList<>();      //для каждой страницы файла прокрутка для таблички
    private List<JButton> readDatas = new ArrayList<>();            //для каждой страницы файла кнопка чтения файла данных
    private List<JLabel> infos = new ArrayList<>();                 //для каждой страницы файла информация о загрузке данных
    private List<JTable> tables = new ArrayList<>();                //для каждой страницы файла таблица с выгруженными данными
    private List<JTable> tablesRez = new ArrayList<>();                //для каждой вкладки с результатом таблица с анализом

    private List<JCheckBox> checkBoxes = new ArrayList<>();         //для каждой страницы файла галочка, работать с листом или нет
    private List<JLabel> labelLoad = new ArrayList<>();             //для каждой страницы информация о загруженных данных

    private List<String[][]> data = new ArrayList<>();               //для каждой строки табличной части файла массив необходимой информации

    private JCheckBox selectINN = new JCheckBox();
    private JCheckBox selectPurpose = new JCheckBox();
    private JCheckBox selectIBalanceSum = new JCheckBox();
    private JCheckBox selectStart = new JCheckBox();
    private JCheckBox selectEnd = new JCheckBox();
    private JCheckBox selectFIO = new JCheckBox();

    private class setPicture implements ActionListener {
        Integer number;
        Integer numberOfSheet;

        public setPicture(Integer number, Integer numberOfSheet) {
            this.number = number;
            this.numberOfSheet = numberOfSheet;
        }

        JLabel pic = new JLabel();

        public void actionPerformed(ActionEvent e) {
            pic = new JLabel(new ImageIcon(getClass().getResource("/" + number + ".png")));


            if (type2.get(numberOfSheet).isSelected()) {
                if (number == 5) {
                    pic = new JLabel(new ImageIcon(getClass().getResource("/7.png")));
                } else if (number == 6) {
                    pic = new JLabel(new ImageIcon(getClass().getResource("/8.png")));
                }
            }

            if (type4.get(numberOfSheet).isSelected()) {
                if (number == 3) {
                    pic = new JLabel(new ImageIcon(getClass().getResource("/10.png")));
                } else if (number == 2) {
                    pic = new JLabel(new ImageIcon(getClass().getResource("/9.png")));
                }
            }

            panelsRight.get(numberOfSheet).remove(examplePanels.get(numberOfSheet));

            example.set(numberOfSheet, pic);

            JPanel panelExample = new JPanel();
            panelExample.setLayout(new BoxLayout(panelExample, BoxLayout.Y_AXIS));

            JPanel panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            JLabel labelTemp = new JLabel();

            if (number == 0) {
                labelTemp = new JLabel("Порядковый номер   ");
            } else if (number == 1) {
                labelTemp = new JLabel("Дата операции   ");
            } else if (number == 2) {
                labelTemp = new JLabel("Сумма прихода (Кт)   ");
            } else if (number == 3) {
                labelTemp = new JLabel("Сумма расхода (Дт)   ");
            } else if (number == 4) {
                labelTemp = new JLabel("Назначение платежа   ");
            } else if (number == 5) {
                labelTemp = new JLabel("Счет получателя (Кт)   ");
            } else if (number == 6) {
                labelTemp = new JLabel("Счет плательщика (Дт)   ");
            }

            if (type2.get(numberOfSheet).isSelected()) {
                if (number == 5) {
                    labelTemp = new JLabel("ИНН контрагента   ");
                } else if (number == 6) {
                    labelTemp = new JLabel("Наименование контрагента   ");
                }
            }

            if (type4.get(numberOfSheet).isSelected()) {
                if (number == 3) {
                    labelTemp = new JLabel("Выписка по счету   ");
                } else if (number == 2) {
                    labelTemp = new JLabel("Сумма операции   ");
                }
            }

            labelTemp.setFont(new Font("Arial", Font.BOLD, 15));

            panel.add(labelTemp);

            JTextField textField = new JTextField();
            textField.setMaximumSize(new Dimension(1000, 30));

            if (number == 0) {
                textField.setText("A");
            } else if (number == 1) {
                textField.setText("C");
            } else if (number == 2) {
                textField.setText("G");
            } else if (number == 3) {
                textField.setText("F");
            } else if (number == 4) {
                textField.setText("H");
            } else if (number == 5) {
                textField.setText("J");
            } else if (number == 6) {
                textField.setText("D");
            }

            if (type2.get(numberOfSheet).isSelected()) {
                if (number == 5) {
                    textField.setText("I");
                } else if (number == 6) {
                    textField.setText("F");
                }
            }

            if (type4.get(numberOfSheet).isSelected()) {
                if (number == 3) {
                    textField.setText("01234567891011121314");
                } else if (number == 2) {
                    textField.setText("Q");
                }
            }
            textField.setEnabled(false);
            textField.setFont(new Font("Arial", Font.BOLD, 16));
            panel.add(textField);
            if (number == 0 || type2.get(numberOfSheet).isSelected() && number == 5) {
                labelTemp = new JLabel(" - необязательное поле");
                labelTemp.setFont(new Font("Arial", Font.BOLD, 15));
                panel.add(labelTemp);
            }
            panel.add(Box.createHorizontalGlue());
            panelExample.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(pic);
            panel.add(Box.createHorizontalGlue());
            panelExample.add(panel);

            examplePanels.set(numberOfSheet, panelExample);

            panelsRight.get(numberOfSheet).add(panelExample);

            repaint();

        }
    }

    //открытие файла
    private class OpenL implements ActionListener {
        public void actionPerformed(ActionEvent e) {

            open.setEnabled(false);

            JFileChooser c = new JFileChooser();
            if (directory.contains("\\")) {
                c = new JFileChooser(directory);
            }
            FileNameExtensionFilter filter = new FileNameExtensionFilter("XLS and XLSX", new String[]{"XLS", "XLSX"});
            c.setFileFilter(filter);


            info.setText("");
            tabbedPane.removeAll();
            label1.setText("");
            label2.setText("");
            label3.setText("");
            label4.setText("");
            label5.setText("");
            label6.setText("");
            label2Info.setText("");
            labelDate.setText("");
            labelOpenExcel.setText("");
            labelSaveData.setText("");
            labelSave1.setText("");
            labelSave2.setText("");
            labelSave3.setText("");
            labelSave4.setText("");
            labelSave5.setText("");
            labelSave6.setText("");

            day.setBackground(Color.WHITE);
            month.setBackground(Color.WHITE);
            year.setBackground(Color.WHITE);
            percent.setBackground(Color.WHITE);
            sum.setBackground(Color.WHITE);
            dayStart.setBackground(Color.WHITE);
            monthStart.setBackground(Color.WHITE);
            yearStart.setBackground(Color.WHITE);
            dayEnd.setBackground(Color.WHITE);
            monthEnd.setBackground(Color.WHITE);
            yearEnd.setBackground(Color.WHITE);

            panelSearch1.remove(scrollPane1);
            JTable jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane1 = new JScrollPane(jTable);
            panelSearch1.add(scrollPane1);

            panelSearch2.remove(scrollPane2);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane2 = new JScrollPane(jTable);
            panelSearch2.add(scrollPane2);

            panelSearch3.remove(scrollPane3);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane3 = new JScrollPane(jTable);
            panelSearch3.add(scrollPane3);

            panelSearch4.remove(scrollPane4);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane4 = new JScrollPane(jTable);
            panelSearch4.add(scrollPane4);

            panelSearch5.remove(scrollPane5);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane5 = new JScrollPane(jTable);
            panelSearch5.add(scrollPane5);

            panelSearch6.remove(scrollPane6);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            scrollPane6 = new JScrollPane(jTable);
            panelSearch6.add(scrollPane6);

            int rVal = c.showOpenDialog(Main.this);
            if (rVal == JFileChooser.APPROVE_OPTION) {
                fileName.setText(c.getSelectedFile().getAbsolutePath());
                name = c.getSelectedFile().getName();
                directory = c.getSelectedFile().getPath();
                if (name.endsWith(".xls")) {
                    name = name.substring(0, name.length() - 4);
                } else {
                    name = name.substring(0, name.length() - 5);
                }
                openExcel.setEnabled(true);
            }
            if (rVal == JFileChooser.CANCEL_OPTION) {
                fileName.setText("Файл не выбран...");
                open.setEnabled(true);
                openExcel.setEnabled(false);

                return;
            }
            if (rVal == JFileChooser.ERROR_OPTION) {
                fileName.setText("Ошибочка при выборе файла(((");
                open.setEnabled(true);
                openExcel.setEnabled(false);
                return;
            }

            readFile();

            open.setEnabled(true);
        }
    }

    //скрытие/показ панели для пользователя постранично
    private class ShowHide implements ActionListener {
        Integer numberOfSheet;

        public ShowHide(Integer numberOfSheet) {
            this.numberOfSheet = numberOfSheet;
        }

        public void actionPerformed(ActionEvent e) {
            if (checkBoxes.get(numberOfSheet).isSelected())
                panelsHide.get(numberOfSheet).setVisible(true);
            else
                panelsHide.get(numberOfSheet).setVisible(false);
        }
    }

    private class Block implements MouseListener {

        public void mousePressed(MouseEvent e) {

            if (jpanelRez6.isVisible()) {
                day.setEnabled(false);
                month.setEnabled(false);
                year.setEnabled(false);
                labelDate.setText("");
            } else {
                day.setEnabled(true);
                month.setEnabled(true);
                year.setEnabled(true);
            }
        }

        public void mouseClicked(MouseEvent e) {
        }

        public void mouseEntered(MouseEvent e) {
        }

        public void mouseExited(MouseEvent e) {
        }

        public void mouseReleased(MouseEvent e) {
        }

    }

    private class SetWhite implements ActionListener {
        List<JTextField> list;
        JLabel label;

        public SetWhite(List<JTextField> list, JLabel label) {
            this.list = list;
            this.label = label;
        }

        public void actionPerformed(ActionEvent e) {
            for (int i = 0; i < list.size(); i++) {
                list.get(i).setBackground(Color.WHITE);
            }
            label.setText("");

        }
    }

    private class Change implements ActionListener {
        Integer numberOfSheet;
        Integer sumType;

        public Change(Integer numberOfSheet, Integer sumType) {
            this.numberOfSheet = numberOfSheet;
            this.sumType = sumType;
        }

        public void actionPerformed(ActionEvent e) {
            switch (sumType) {
                case (0):
                    if (type1.get(numberOfSheet).isSelected()) {
                        INNNumsPlus.get(numberOfSheet).setText("  счет получателя(Кт)");
                        INNNumsMinus.get(numberOfSheet).setText("  счет плательщика(Дт)");

                        showPic("/DK.png", "Выписка содержит отдельно реквизиты счетов Дт и Кт:", numberOfSheet);
                    } else {
                        INNNumsPlus.get(numberOfSheet).setText("  ИНН/счет контрагента");
                        INNNumsMinus.get(numberOfSheet).setText("  наим-ие контрагента");

                        //type3.get(numberOfSheet).setSelected(true);

                        showPic("/INN.png", "Выписка содержит реквизиты контрагента:", numberOfSheet);
                    }
                    break;
                case (1):
                    if (type3.get(numberOfSheet).isSelected()) {
                        sumNumsPlus.get(numberOfSheet).setText("  сумма прихода (Кт)");
                        sumNumsMinus.get(numberOfSheet).setText("  сумма расхода (Дт)");
                        jsumNumMinus.get(numberOfSheet).setColumns(2);
                        jsumNumMinus.get(numberOfSheet).setMaximumSize(new Dimension(10000, 40));

                        showPic("/DKSum.png", "Выписка содержит отдельно суммы по Дт и Кт:", numberOfSheet);
                    } else {
                        sumNumsPlus.get(numberOfSheet).setText("  сумма операции");
                        sumNumsMinus.get(numberOfSheet).setText("  выписка по счету");
                        jsumNumMinus.get(numberOfSheet).setColumns(17);
                        jsumNumMinus.get(numberOfSheet).setMaximumSize(new Dimension(100000, 40));

                        //type1.get(numberOfSheet).setSelected(true);

                        showPic("/Sum.png", "Выписка содержит одну колонку с суммой:", numberOfSheet);
                    }
                    break;
            }
        }


    }

    public void showPic(String file, String text, int numberOfSheet) {
        JLabel pic = new JLabel(new ImageIcon(getClass().getResource(file)));

        panelsRight.get(numberOfSheet).remove(examplePanels.get(numberOfSheet));

        example.set(numberOfSheet, pic);

        JPanel picPanel = new JPanel();
        picPanel.setLayout(new BoxLayout(picPanel, BoxLayout.Y_AXIS));
        JLabel name = new JLabel(text);
        name.setFont(new Font("Arial", Font.BOLD, 15));
        picPanel.add(name);
        picPanel.add(Box.createRigidArea(new Dimension(0, 10)));
        picPanel.add(pic);
        examplePanels.set(numberOfSheet, picPanel);
        panelsRight.get(numberOfSheet).add(picPanel);

        repaint();

    }

    //чтение файла
    public void readFile() {
        int numberOfSheets = 0;
        try {

            FileInputStream file = new FileInputStream(new File(fileName.getText()));

            if (fileName.getText().endsWith(".xls")) {
                hssfWorkbook = new HSSFWorkbook(file);
                numberOfSheets = hssfWorkbook.getNumberOfSheets();
            } else if (fileName.getText().endsWith(".xlsx")) {
                xssfWorkbook = new XSSFWorkbook(file);
                numberOfSheets = xssfWorkbook.getNumberOfSheets();
            }
            info.setText("Листов в файле: " + numberOfSheets);
        } catch (Exception ee) {
            info.setText("Проблемы с открытием файла(((");
        }
        createSheets(numberOfSheets);
    }

    //создание листов для заполнения пользователем
    public void createSheets(int numberOfSheets) {

        panelsHide.clear();
        progressBars.clear();
        readDatas.clear();
        infos.clear();
        tables.clear();
        checkBoxes.clear();
        panelsDown.clear();
        jrowStarts.clear();
        jnums.clear();
        jdateNums.clear();
        jINNNumsPlus.clear();
        jINNNumsMinus.clear();
        jpurposeNums.clear();
        jsumNumPlus.clear();
        jsumNumMinus.clear();
        data.clear();
        scrollPanes.clear();
        panelsRightTable.clear();
        labelLoad.clear();
        type1.clear();
        type2.clear();
        INNNumsMinus.clear();
        INNNumsPlus.clear();
        examplePanels.clear();
        example.clear();
        panelsRight.clear();
        type3.clear();
        type4.clear();
        sumNumsMinus.clear();
        sumNumsPlus.clear();
        tablesRez.clear();
        JTable jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        for (int i = 0; i < 6; i++) {
            tablesRez.add(jTable);
        }

        tabbedPane.removeAll();

        jpanel.add(tabbedPane);
        for (int i = 0; i < numberOfSheets; i++) {
            JPanel panelTabbedPane = new JPanel();
            tabbedPane.addTab("Лист " + (i + 1), panelTabbedPane);
            panelTabbedPane.setLayout(new BoxLayout(panelTabbedPane, BoxLayout.Y_AXIS));
////////////////////////////////////////////////////////////////////////////////////////////////////
            JPanel panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            panelTabbedPane.add(panel);
            panelTabbedPane.add(Box.createRigidArea(new Dimension(0, 10)));
////////////////////////////////////////////////////////////////////////////////////////////////////
            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            JCheckBox checkBox = new JCheckBox();
            checkBox.setText("Данный лист содержит сделки для анализа");
            checkBox.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            checkBox.addActionListener(new ShowHide(i));
            checkBoxes.add(checkBox);
            panel.add(checkBox);
            panel.add(Box.createHorizontalGlue());
            panelTabbedPane.add(panel);
            panelTabbedPane.add(Box.createRigidArea(new Dimension(0, 10)));
////////////////////////////////////////////////////////////////////////////////////////////////////
            JPanel panelHide = new JPanel();
            panelHide.setLayout(new BoxLayout(panelHide, BoxLayout.Y_AXIS));
            Border etched = BorderFactory.createEtchedBorder();
            panelHide.setBorder(etched);
            panelsHide.add(panelHide);

            JPanel panelUp = new JPanel();
            panelUp.setLayout(new GridLayout(1, 2));
            panelUp.setBorder(etched);

            JPanel panelDown = new JPanel();
            panelDown.setLayout(new BoxLayout(panelDown, BoxLayout.Y_AXIS));
            panelDown.setBorder(etched);
            panelsDown.add(panelDown);

            JPanel panelRight = new JPanel();
            panelRight.setBorder(etched);
            panelsRight.add(panelRight);

            JPanel panelLeft = new JPanel();
            panelLeft.setLayout(new BoxLayout(panelLeft, BoxLayout.Y_AXIS));
            panelLeft.setBorder(etched);

///////////////////////////////////////////////////////////////////////////////////////////////////////

            JRadioButton b1 = new JRadioButton();
            b1.setText("Выписка содержит отдельно реквизиты счетов Дт и Кт  ");
            b1.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            type1.add(b1);
            JRadioButton b2 = new JRadioButton();
            b2.setText("Выписка содержит реквизиты контрагента (в разных колонках)  ");
            b2.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            type2.add(b2);
            b1.addActionListener(new Change(i, 0));
            b2.addActionListener(new Change(i, 0));

            b1.setSelected(true);

            ButtonGroup gr1 = new ButtonGroup();
            gr1.add(b1);
            gr1.add(b2);

            JRadioButton b3 = new JRadioButton();
            b3.setText("Выписка содержит отдельно суммы по Дт и Кт  ");
            b3.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            type3.add(b3);
            JRadioButton b4 = new JRadioButton();
            b4.setText("Выписка содержит одну колонку с суммой  ");
            b4.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            type4.add(b4);
            b3.addActionListener(new Change(i, 1));
            b4.addActionListener(new Change(i, 1));

            b3.setSelected(true);

            ButtonGroup gr2 = new ButtonGroup();
            gr2.add(b3);
            gr2.add(b4);

            panel = new JPanel();
            panel.setLayout(new GridLayout(2, 2));

            panel.add(b1);
            panel.add(b3);
            panel.add(b2);
            panel.add(b4);

            panelHide.add(panel);
            panelHide.add(Box.createRigidArea(new Dimension(0, 10)));

            ///////////////////////////////////////////////////////////////////////////////////////////////
            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            panelLeft.add(panel);
            panelLeft.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////////
            JPanel panelInput = new JPanel();
            panelInput.setLayout(new GridLayout(7, 2));

            panelInput.setBorder(BorderFactory.createTitledBorder("Введите названия колонок Excel, в которых находятся:"));
            ((javax.swing.border.TitledBorder) panelInput.getBorder()).
                    setTitleFont(new Font("Arial", Font.BOLD, 14));
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            JLabel labelTemp = new JLabel("  порядковый номер");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            JTextField textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jnums.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[0] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[0].addActionListener(new setPicture(0, i));
            helps[0].setToolTipText("Показать пример заполнения");
            panel.add(helps[0]);
            panelInput.add(panel);

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  дата операции");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

            panelInput.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jdateNums.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[1] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[1].addActionListener(new setPicture(1, i));
            helps[1].setToolTipText("Показать пример заполнения");
            panel.add(helps[1]);
            panelInput.add(panel);
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  сумма расхода (Дт)");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);
            sumNumsMinus.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jsumNumMinus.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[3] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[3].addActionListener(new setPicture(3, i));
            helps[3].setToolTipText("Показать пример заполнения");
            panel.add(helps[3]);
            panelInput.add(panel);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  сумма прихода (Кт)");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);
            sumNumsPlus.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jsumNumPlus.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[2] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[2].addActionListener(new setPicture(2, i));
            helps[2].setToolTipText("Показать пример заполнения");
            panel.add(helps[2]);
            panelInput.add(panel);

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  назначение платежа");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jpurposeNums.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[4] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[4].addActionListener(new setPicture(4, i));
            helps[4].setToolTipText("Показать пример заполнения");
            panel.add(helps[4]);
            panelInput.add(panel);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  счет плательщика(Дт)");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);
            INNNumsMinus.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jINNNumsMinus.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[6] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[6].addActionListener(new setPicture(6, i));
            helps[6].setToolTipText("Показать пример заполнения");
            panel.add(helps[6]);
            panelInput.add(panel);
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            labelTemp = new JLabel("  счет получателя(Кт)");
            labelTemp.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            panelInput.add(labelTemp);
            INNNumsPlus.add(labelTemp);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(Box.createHorizontalGlue());
            textField = new JTextField();
            textField.setColumns(2);
            textField.setMaximumSize(new Dimension(10000, 40));
            textField.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jINNNumsPlus.add(textField);
            panel.add(textField);
            panelInput.add(panel);

            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            panel.add(new JLabel(""));
            panel.add(Box.createHorizontalGlue());
            helps[5] = new JButton(new ImageIcon(getClass().getResource("/16.jpg")));
            helps[5].addActionListener(new setPicture(5, i));
            helps[5].setToolTipText("Показать пример заполнения");
            panel.add(helps[5]);
            panelInput.add(panel);

            panelLeft.add(panelInput);
            panelUp.add(panelLeft);
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            JLabel pic = new JLabel(new ImageIcon(getClass().getResource("/DK.png")));
            example.add(pic);

            JPanel picPanel = new JPanel();
            picPanel.setLayout(new BoxLayout(picPanel, BoxLayout.Y_AXIS));
            JLabel text = new JLabel("Выписка содержит отдельно реквизиты счетов Дт и Кт:");
            text.setFont(new Font("Arial", Font.BOLD, 15));
            picPanel.add(text);
            picPanel.add(Box.createRigidArea(new Dimension(0, 10)));
            picPanel.add(pic);
            examplePanels.add(picPanel);
            panelRight.add(picPanel);

            panelUp.add(panelRight);
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panelHide.add(panelUp);
            panelHide.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            JButton readData = new JButton();
            readData.setText("Получить данные из файла");
            readData.setFont(new Font("Arial", Font.BOLD, 14));
            readData.addActionListener(new ReadData(i));
            readDatas.add(readData);
            panel.add(readData);
            JLabel info = new JLabel();
            info.setFont(new Font("Arial", Font.BOLD, 14));
            infos.add(info);
            panel.add(info);
            panel.add(Box.createHorizontalGlue());
            panelDown.add(panel);
            panelDown.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
            panelsRightTable.add(panel);
            jTable = new JTable(tableData, headers);
            jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jTable.setRowHeight(20);
            tables.add(jTable);
            JScrollPane scrollPane = new JScrollPane(jTable);
            scrollPanes.add(scrollPane);
            panel.add(scrollPane);
            panelDown.add(panel);
            panelDown.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
            info = new JLabel();
            info.setFont(new Font("Arial", Font.BOLD, 14));
            labelLoad.add(info);
            panel.add(info);
            panel.add(Box.createHorizontalGlue());
            panelDown.add(panel);
            panelDown.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panelDown.add(Box.createHorizontalGlue());
            panelHide.add(panelDown);
            panelHide.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////////////////
            panelHide.add(Box.createHorizontalGlue());
            panelTabbedPane.add(panelHide);
            panelTabbedPane.add(Box.createRigidArea(new Dimension(0, 10)));

            panelHide.setVisible(false);
            String[][] tableData = {{"", "", "", "", "", "", ""}};
            data.add(tableData);
        }
    }

    //загрузка данных файла
    private class ReadData implements ActionListener {

        Integer numberOfSheet;

        public ReadData(Integer numberOfSheet) {
            this.numberOfSheet = numberOfSheet;
        }

        public void actionPerformed(ActionEvent e) {

            boolean ok = true;

            int num = 0;
            try {
                String key = jnums.get(numberOfSheet).getText().replaceAll(" ", "");
                num = map.get(key);
                jnums.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {
            }

            int dateNum = 0;
            try {
                String key = jdateNums.get(numberOfSheet).getText().replaceAll(" ", "");
                dateNum = map.get(key);
                jdateNums.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {
                jdateNums.get(numberOfSheet).setBackground(Color.RED);
                ok = false;
            }

            int INNNumPlus = 0;
            try {
                String key = jINNNumsPlus.get(numberOfSheet).getText().replaceAll(" ", "");
                INNNumPlus = map.get(key);
                jINNNumsPlus.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {
                if (type1.get(numberOfSheet).isSelected()) {
                    jINNNumsPlus.get(numberOfSheet).setBackground(Color.RED);
                    ok = false;
                } else
                    jINNNumsPlus.get(numberOfSheet).setBackground(Color.WHITE);
            }

            int INNNumMinus = 0;
            try {
                String key = jINNNumsMinus.get(numberOfSheet).getText().replaceAll(" ", "");
                INNNumMinus = map.get(key);
                jINNNumsMinus.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {
                jINNNumsMinus.get(numberOfSheet).setBackground(Color.RED);
                ok = false;
            }

            int sumNumPlus = 0;
            try {
                String key = jsumNumPlus.get(numberOfSheet).getText().replaceAll(" ", "");
                sumNumPlus = map.get(key);
                jsumNumPlus.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {
                jsumNumPlus.get(numberOfSheet).setBackground(Color.RED);
                ok = false;
            }

            int sumNumMinus = 0;
            String bill = "";
            if (type3.get(numberOfSheet).isSelected()) {
                try {
                    String key = jsumNumMinus.get(numberOfSheet).getText().replaceAll(" ", "");
                    sumNumMinus = map.get(key);
                    jsumNumMinus.get(numberOfSheet).setBackground(Color.WHITE);
                } catch (Exception ex) {
                    jsumNumMinus.get(numberOfSheet).setBackground(Color.RED);
                    ok = false;
                }
            } else if (type4.get(numberOfSheet).isSelected()) {
                bill = jsumNumMinus.get(numberOfSheet).getText().replaceAll(" ", "");
                if (bill.equals("")) {
                    ok = false;
                } else {
                    jsumNumMinus.get(numberOfSheet).setBackground(Color.WHITE);
                }
            }

            int purposeNum = 0;
            try {
                String key = jpurposeNums.get(numberOfSheet).getText().replaceAll(" ", "");
                purposeNum = map.get(key);
                jpurposeNums.get(numberOfSheet).setBackground(Color.WHITE);
            } catch (Exception ex) {

                jpurposeNums.get(numberOfSheet).setBackground(Color.RED);
                ok = false;
            }

            if (!ok) {
                infos.get(numberOfSheet).setText(" Некорректно введены данные");
                infos.get(numberOfSheet).setForeground(Color.RED);
                return;
            }
            infos.get(numberOfSheet).setText("");

            String[][] forTable;
            List<String[]> list = new ArrayList<>();

            if (fileName.getText().endsWith(".xls")) {
                HSSFSheet sheet = hssfWorkbook.getSheetAt(numberOfSheet);

                for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                    HSSFRow row = sheet.getRow(j);

                    try {
                        if (row.getCell(dateNum - 1) != null) {

                            String[] array = new String[9];

                            if (num == 0 || row.getCell(num - 1) == null) {
                                array[0] = "";
                            } else array[0] = row.getCell(num - 1).toString();

                            if (row.getCell(purposeNum - 1) == null) {
                                array[6] = "";
                            } else array[6] = row.getCell(purposeNum - 1).toString();


                            if (type3.get(numberOfSheet).isSelected()) {
                                if (row.getCell(sumNumPlus - 1) == null) {
                                    array[4] = "";
                                } else {
                                    if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                        array[4] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                    } else array[4] = row.getCell(sumNumPlus - 1).toString();
                                }

                                if (row.getCell(sumNumMinus - 1) == null) {
                                    array[5] = "";
                                } else {
                                    if (row.getCell(sumNumMinus - 1).getCellType() == CellType.NUMERIC) {
                                        array[5] = nf.format(row.getCell(sumNumMinus - 1).getNumericCellValue());
                                    } else array[5] = row.getCell(sumNumMinus - 1).toString();
                                }
                            } else if (type4.get(numberOfSheet).isSelected()) {

                                if (row.getCell(INNNumMinus - 1).toString().contains(bill)) {
                                    if (row.getCell(sumNumPlus - 1) == null) {
                                        array[5] = "";
                                    } else {
                                        if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                            array[5] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                        } else array[5] = row.getCell(sumNumPlus - 1).toString();
                                    }

                                    array[4] = "";
                                } else if (row.getCell(INNNumPlus - 1).toString().contains(bill)) {
                                    if (row.getCell(sumNumPlus - 1) == null) {
                                        array[4] = "";
                                    } else {
                                        if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                            array[4] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                        } else array[4] = row.getCell(sumNumPlus - 1).toString();
                                    }
                                    array[5] = "";
                                }
                            }


                            if (type1.get(numberOfSheet).isSelected()) {
                                if (findSum(array[4])) {
                                    if (row.getCell(INNNumMinus - 1) == null) {
                                        array[2] = "";
                                        array[3] = "";
                                    } else {
                                        array[2] = getINN(row.getCell(INNNumMinus - 1).toString());
                                        array[3] = getFIO(row.getCell(INNNumMinus - 1).toString());
                                    }
                                } else if (findSum(array[5])) {
                                    if (row.getCell(INNNumPlus - 1) == null) {
                                        array[2] = "";
                                        array[3] = "";
                                    } else {
                                        array[2] = getINN(row.getCell(INNNumPlus - 1).toString());
                                        array[3] = getFIO(row.getCell(INNNumPlus - 1).toString());
                                    }
                                }
                            } else if (type2.get(numberOfSheet).isSelected()) {
                                if (INNNumPlus == 0) {
                                    array[2] = "";
                                } else if (row.getCell(INNNumPlus - 1) == null) {
                                    array[2] = "";
                                } else {
                                    if (row.getCell(INNNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                        array[2] = nf.format(row.getCell(INNNumPlus - 1).getNumericCellValue());
                                    } else array[2] = row.getCell(INNNumPlus - 1).toString().replaceAll("\\D+", "");
                                }
                            if (row.getCell(INNNumMinus - 1) == null) {
                                array[3] = "";
                            } else {
                                array[3] = row.getCell(INNNumMinus - 1).toString();
                            }
                        }

                        if (!(array[4].equals("") && (array[5].equals("")))) {
                            if (findDate(row.getCell(dateNum - 1).toString())) {
                                array[1] = row.getCell(dateNum - 1).toString();
                                list.add(array);
                            } else if (!row.getCell(dateNum - 1).toString().equals("" + row.getCell(dateNum - 1).getNumericCellValue())) {
                                array[1] = new SimpleDateFormat("dd.MM.yyyy").format(row.getCell(dateNum - 1).getDateCellValue());
                                list.add(array);
                            }
                        }

                    }
                } catch(Exception ee){
                }
            }
        } else

        {
            XSSFSheet sheet = xssfWorkbook.getSheetAt(numberOfSheet);

            for (int j = 0; j <= sheet.getLastRowNum(); j++) {
                XSSFRow row = sheet.getRow(j);

                try {
                    if (row.getCell(dateNum - 1) != null) {

                        String[] array = new String[9];

                        if (num == 0 || row.getCell(num - 1) == null) {
                            array[0] = "";
                        } else array[0] = row.getCell(num - 1).toString();

                        if (type3.get(numberOfSheet).isSelected()) {
                            if (row.getCell(sumNumPlus - 1) == null) {
                                array[4] = "";
                            } else {
                                if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                    array[4] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                } else array[4] = row.getCell(sumNumPlus - 1).toString();
                            }

                            if (row.getCell(sumNumMinus - 1) == null) {
                                array[5] = "";
                            } else {
                                if (row.getCell(sumNumMinus - 1).getCellType() == CellType.NUMERIC) {
                                    array[5] = nf.format(row.getCell(sumNumMinus - 1).getNumericCellValue());
                                } else array[5] = row.getCell(sumNumMinus - 1).toString();
                            }
                        } else if (type4.get(numberOfSheet).isSelected()) {
                            if (row.getCell(INNNumMinus - 1).toString().contains(bill)) {
                                if (row.getCell(sumNumPlus - 1) == null) {
                                    array[5] = "";
                                } else {
                                    if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                        array[5] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                    } else array[5] = row.getCell(sumNumPlus - 1).toString();
                                }
                                array[4] = "";
                            } else if (row.getCell(INNNumPlus - 1).toString().contains(bill)) {
                                if (row.getCell(sumNumPlus - 1) == null) {
                                    array[4] = "";
                                } else {
                                    if (row.getCell(sumNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                        array[4] = nf.format(row.getCell(sumNumPlus - 1).getNumericCellValue());
                                    } else array[4] = row.getCell(sumNumPlus - 1).toString();
                                }
                                array[5] = "";
                            }
                        }

                        if (row.getCell(purposeNum - 1) == null) {
                            array[6] = "";
                        } else array[6] = row.getCell(purposeNum - 1).toString();

                        if (type1.get(numberOfSheet).isSelected()) {
                            if (findSum(array[4])) {
                                if (row.getCell(INNNumMinus - 1) == null) {
                                    array[2] = "";
                                    array[3] = "";
                                } else {
                                    array[2] = getINN(row.getCell(INNNumMinus - 1).toString());
                                    array[3] = getFIO(row.getCell(INNNumMinus - 1).toString());
                                }
                            } else if (findSum(array[5])) {
                                if (row.getCell(INNNumPlus - 1) == null) {
                                    array[2] = "";
                                    array[3] = "";
                                } else {
                                    array[2] = getINN(row.getCell(INNNumPlus - 1).toString());
                                    array[3] = getFIO(row.getCell(INNNumPlus - 1).toString());
                                }
                            }
                        } else if (type2.get(numberOfSheet).isSelected()) {
                            if (INNNumPlus == 0) {
                                array[2] = "";
                            } else if (row.getCell(INNNumPlus - 1) == null) {
                                array[2] = "";
                            } else {
                                if (row.getCell(INNNumPlus - 1).getCellType() == CellType.NUMERIC) {
                                    array[2] = nf.format(row.getCell(INNNumPlus - 1).getNumericCellValue());
                                } else array[2] = row.getCell(INNNumPlus - 1).toString().replaceAll("\\D+", "");
                            }
                            if (row.getCell(INNNumMinus - 1) == null) {
                                array[3] = "";
                            } else {
                                array[3] = row.getCell(INNNumMinus - 1).toString();
                            }
                        }

                        if (!(array[4].equals("") && (array[5].equals("")))) {
                            if (findDate(row.getCell(dateNum - 1).toString())) {
                                array[1] = row.getCell(dateNum - 1).toString();
                                list.add(array);
                            } else if (!row.getCell(dateNum - 1).toString().equals("" + row.getCell(dateNum - 1).getNumericCellValue())) {
                                array[1] = new SimpleDateFormat("dd.MM.yyyy").format(row.getCell(dateNum - 1).getDateCellValue());
                                list.add(array);
                            }
                        }
                    }
                } catch (Exception ee) {
                }
            }
        }

        forTable =new String[list.size()][9];
            for(
        int i = 0; i<list.size();i++)

        {
            forTable[i] = list.get(i);
        }

            data.set(numberOfSheet,forTable);

            panelsRightTable.get(numberOfSheet).

        remove(scrollPanes.get(numberOfSheet));

        repaint();

        JTable jTable = new JTable(forTable, headers);
            jTable.getTableHeader().

        setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
            jTable.setFont(new

        Font("Arial",Font.TRUETYPE_FONT, 13));
            tables.set(numberOfSheet,jTable);

        JScrollPane scrollPane = new JScrollPane(jTable);
            scrollPanes.set(numberOfSheet,scrollPane);

            panelsRightTable.get(numberOfSheet).

        add(scrollPane);

        setWidth(jTable);

            labelLoad.get(numberOfSheet).

        setText("Загружено сделок: "+list.size() +
                ". Сумма операций (Кт): "+

        getSum(list, 4) +
                ". Сумма операций (Дт): "+

        getSum(list, 5));

    }
}

    public void setWidth(JTable jTable) {
        TableColumn column = jTable.getColumnModel().getColumn(0);
        column.setMaxWidth(40);
        column = jTable.getColumnModel().getColumn(1);
        column.setMaxWidth(85);
        column = jTable.getColumnModel().getColumn(2);
        column.setPreferredWidth(100);
        column.setMaxWidth(200);
        column = jTable.getColumnModel().getColumn(3);
        column.setPreferredWidth(150);
        column.setMaxWidth(500);
        column = jTable.getColumnModel().getColumn(4);
        column.setMaxWidth(95);
        column = jTable.getColumnModel().getColumn(5);
        column.setMaxWidth(95);
    }

    public String getINN(String text) {
        String[] array = text.split("\n");
        if (array.length == 3 && array[1].matches("\\d+"))
            return array[1];
        else return "";
    }

    public String getFIO(String text) {
        String[] array = text.split("\n");
        if (array.length == 3 && array[1].matches("\\d+"))
            return array[2];
        else return text;
    }

    public String getSum(List<String[]> list, int num) {

        double sum = 0.0;

        for (int i = 0; i < list.size(); i++) {
            String INNSumString = list.get(i)[num].replaceAll(" ", "")
                    .replaceAll("\\u00A0", "")
                    .replaceAll("-", ".")
                    .replaceAll(",", ".");
            try {
                if (findSum(INNSumString))
                    sum += Double.parseDouble(INNSumString);
            } catch (Exception e) {
            }
        }
        return NumberFormat.getInstance(new Locale("ru", "RU")).format(sum);
    }

private class search implements ActionListener {

    private int numberOfButton;

    public search(int numberOfButton) {
        this.numberOfButton = numberOfButton;
    }

    public void actionPerformed(ActionEvent e) {

        int dayB = 0;
        int monthB = 0;
        int yearB = 0;

        boolean ok = true;

        try {
            dayB = Integer.parseInt(day.getText().replaceAll(" ", ""));
            day.setBackground(Color.WHITE);
        } catch (Exception ee) {
            day.setBackground(Color.RED);
            ok = false;
        }

        try {
            monthB = Integer.parseInt(month.getText().replaceAll(" ", ""));
            month.setBackground(Color.WHITE);
        } catch (Exception ee) {
            month.setBackground(Color.RED);
            ok = false;
        }

        try {
            yearB = Integer.parseInt(year.getText().replaceAll(" ", ""));
            year.setBackground(Color.WHITE);
        } catch (Exception ee) {
            year.setBackground(Color.RED);
            ok = false;
        }

        if (!ok) {
            labelDate.setText("   Дата принятия заявления о банкротстве задана некорректно");
            labelDate.setForeground(Color.RED);
            return;
        }
        labelDate.setText("");

        List<String[]> list = new ArrayList<>();
        List<String> listINN;
        List<String> words;

        if (selectINN.isSelected()) {
            listINN = getList(areaINN, 0);
        } else {
            listINN = null;
        }

        if (selectPurpose.isSelected()) {
            words = getList(areaPurpose, 1);
        } else {
            words = null;
        }

        for (int i = 0; i < data.size(); i++) {
            if (checkBoxes.get(i).isSelected()) {
                for (int j = 0; j < data.get(i).length; j++) {

                    if (numberOfButton == 1) {
                        if (findSum(data.get(i)[j][5])) {
                            if (checkPurpose(data.get(i)[j][6], dayB, monthB, yearB) && checkDate(data.get(i)[j][1], dayB, monthB, yearB, 1, 0)) {
                                if (isSelect(listINN, i, j, words)) list.add(data.get(i)[j]);
                            }
                        }
                    } else if (numberOfButton == 2) {
                        if (findSum(data.get(i)[j][5])) {
                            if (checkPurpose(data.get(i)[j][6], dayB, monthB, yearB) && checkDate(data.get(i)[j][1], dayB, monthB, yearB, 6, 0)) {
                                if (isSelect(listINN, i, j, words)) list.add(data.get(i)[j]);
                            }
                        }
                    } else if (numberOfButton == 3) {
                        if (checkDate(data.get(i)[j][1], dayB, monthB, yearB, 0, 1)) {
                            if (isSelect(listINN, i, j, words)) list.add(data.get(i)[j]);
                        }
                    } else if (numberOfButton == 4) {
                        if (checkDate(data.get(i)[j][1], dayB, monthB, yearB, 0, 3)) {
                            if (isSelect(listINN, i, j, words)) list.add(data.get(i)[j]);
                        }
                    } else if (numberOfButton == 5) {
                        if (checkDate(data.get(i)[j][1], dayB, monthB, yearB, 0, 10)) {
                            if (isSelect(listINN, i, j, words)) list.add(data.get(i)[j]);
                        }
                    }
                }
            }
        }
        List<String> listINNforSum = new ArrayList<>();

        if (selectIBalanceSum.isSelected()) {

            String balanceSum = sum.getText().replaceAll(" ", "");

            List<String[]> listCopy = new ArrayList<>();

            listINNforSum = getListINNforSum(balanceSum, list, 2, 5);

            if (listINNforSum.size() > 0) {
                for (int i = 0; i < list.size(); i++) {
                    if (listINNforSum.contains(list.get(i)[2]) && findSum(list.get(i)[5]))
                        listCopy.add(list.get(i));
                }
            }

            listINNforSum = getListINNforSum(balanceSum, list, 2, 4);

            if (listINNforSum.size() > 0) {
                for (int i = 0; i < list.size(); i++) {
                    if (listINNforSum.contains(list.get(i)[2]) && findSum(list.get(i)[4]))
                        listCopy.add(list.get(i));
                }
            }
            list.clear();
            list.addAll(listCopy);
        }

        Object[][] first = new Object[list.size()][7];

        for (int i = 0; i < list.size(); i++) {
            first[i] = list.get(i);
        }

        String text = "  Найдено сделок:" + list.size() +
                ". Сумма операций (Кт): " + getSum(list, 4) +
                ". Сумма операций (Дт): " + getSum(list, 5);

        JTable jTable = new JTable(first, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        jTable.setFont(new Font("Arial", Font.TRUETYPE_FONT, 13));
        tablesRez.set(numberOfButton - 1, jTable);
        setWidth(jTable);

        switch (numberOfButton) {
            case 1:
                label1.setText(text);
                panelSearch1.remove(scrollPane1);
                repaint();
                scrollPane1 = new JScrollPane(jTable);
                panelSearch1.add(scrollPane1);
                break;
            case 2:
                label2.setText(text);
                panelSearch2.remove(scrollPane2);
                repaint();
                scrollPane2 = new JScrollPane(jTable);
                panelSearch2.add(scrollPane2);
                break;
            case 3:
                label3.setText(text);
                panelSearch3.remove(scrollPane3);
                repaint();
                scrollPane3 = new JScrollPane(jTable);
                panelSearch3.add(scrollPane3);
                break;
            case 4:
                label4.setText(text);
                panelSearch4.remove(scrollPane4);
                repaint();
                scrollPane4 = new JScrollPane(jTable);
                panelSearch4.add(scrollPane4);
                break;
            case 5:
                label5.setText(text);
                panelSearch5.remove(scrollPane5);
                repaint();
                scrollPane5 = new JScrollPane(jTable);
                panelSearch5.add(scrollPane5);
                break;
        }
    }
}

private class searchAll implements ActionListener {

    public void actionPerformed(ActionEvent e) {

        int dayS = 0;
        int monthS = 0;
        int yearS = 0;
        int dayF = 0;
        int monthF = 0;
        int yearF = 0;

        boolean ok = true;

        if (selectStart.isSelected()) {
            try {
                dayS = Integer.parseInt(dayStart.getText().replaceAll(" ", ""));
                dayStart.setBackground(Color.WHITE);
            } catch (Exception ee) {
                dayStart.setBackground(Color.RED);
                ok = false;
            }

            try {
                monthS = Integer.parseInt(monthStart.getText().replaceAll(" ", ""));
                monthStart.setBackground(Color.WHITE);
            } catch (Exception ee) {
                monthStart.setBackground(Color.RED);
                ok = false;
            }

            try {
                yearS = Integer.parseInt(yearStart.getText().replaceAll(" ", ""));
                yearStart.setBackground(Color.WHITE);
            } catch (Exception ee) {
                yearStart.setBackground(Color.RED);
                ok = false;
            }
        }

        if (!ok) {
            label6.setText("   Дата начала периода задана некорректно");
            label6.setForeground(Color.RED);
            return;
        }
        label6.setText("");

        if (selectEnd.isSelected()) {
            try {
                dayF = Integer.parseInt(dayEnd.getText().replaceAll(" ", ""));
                dayEnd.setBackground(Color.WHITE);
            } catch (Exception ee) {
                dayEnd.setBackground(Color.RED);
                ok = false;
            }

            try {
                monthF = Integer.parseInt(monthEnd.getText().replaceAll(" ", ""));
                monthEnd.setBackground(Color.WHITE);
            } catch (Exception ee) {
                monthEnd.setBackground(Color.RED);
                ok = false;
            }

            try {
                yearF = Integer.parseInt(yearEnd.getText().replaceAll(" ", ""));
                yearEnd.setBackground(Color.WHITE);
            } catch (Exception ee) {
                yearEnd.setBackground(Color.RED);
                ok = false;
            }
        }

        if (!ok) {
            label6.setText("   Дата окончания периода задана некорректно");
            label6.setForeground(Color.RED);
            return;
        }
        label6.setText("");

        List<String[]> list = new ArrayList<>();
        List<String> listINN;
        List<String> words;
        List<String> FIO;

        if (selectINN.isSelected()) {
            listINN = getList(areaINN, 0);
        } else {
            listINN = null;
        }

        if (selectPurpose.isSelected()) {
            words = getList(areaPurpose, 1);
        } else {
            words = null;
        }

        if (selectFIO.isSelected()) {
            FIO = getList(areaFIO, 1);
        } else {
            FIO = null;
        }

        for (int i = 0; i < data.size(); i++) {
            if (checkBoxes.get(i).isSelected()) {
                for (int j = 0; j < data.get(i).length; j++) {
                    if (checkInterval(data.get(i)[j][1], dayS, monthS, yearS, dayF, monthF, yearF)) {
                        if (isSelect(listINN, i, j, words))
                            if (FIO != null && (findWords(FIO, data.get(i)[j][3]))) {
                                list.add(data.get(i)[j]);
                            } else if (FIO == null) {
                                list.add(data.get(i)[j]);
                            }
                    }
                }
            }
        }
        List<String> listINNforSum = new ArrayList<>();

        if (selectIBalanceSum.isSelected()) {

            String balanceSum = sum.getText().replaceAll(" ", "");

            List<String[]> listCopy = new ArrayList<>();

            listINNforSum = getListINNforSum(balanceSum, list, 2, 5);

            if (listINNforSum.size() > 0) {
                for (int i = 0; i < list.size(); i++) {
                    if (listINNforSum.contains(list.get(i)[2]) && findSum(list.get(i)[5]))
                        listCopy.add(list.get(i));
                }
            }

            listINNforSum = getListINNforSum(balanceSum, list, 2, 4);

            if (listINNforSum.size() > 0) {
                for (int i = 0; i < list.size(); i++) {
                    if (listINNforSum.contains(list.get(i)[2]) && findSum(list.get(i)[4]))
                        listCopy.add(list.get(i));
                }
            }
            list.clear();
            list.addAll(listCopy);
        }

        Object[][] first = new Object[list.size()][7];

        for (int i = 0; i < list.size(); i++) {
            first[i] = list.get(i);
        }

        String text = "  Найдено сделок:" + list.size() +
                ". Сумма операций (Кт): " + getSum(list, 4) +
                ". Сумма операций (Дт): " + getSum(list, 5);

        JTable jTable = new JTable(first, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        jTable.setFont(new Font("Arial", Font.TRUETYPE_FONT, 13));
        tablesRez.set(5, jTable);
        setWidth(jTable);

        label6.setText(text);
        label6.setForeground(Color.BLACK);
        panelSearch6.remove(scrollPane6);
        repaint();
        scrollPane6 = new JScrollPane(jTable);
        panelSearch6.add(scrollPane6);
    }
}

    public boolean checkInterval(String date, int dayS, int monthS, int yearS, int dayF, int monthF, int yearF) {
        String[] dateRow = date.replaceAll(" ", "").split("\\.");
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("/");
        }
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("-");
        }
        if (dateRow.length < 3) {
            labelDate.setText("   Дата операции не распозналась (неизвестный разделитель). Отбор по дате не установлен.");
            return true;
        }

        int dayRow = 0;
        int monthRow = 0;
        int yearRow = 0;

        try {
            dayRow = Integer.parseInt(dateRow[0]);
            monthRow = Integer.parseInt(dateRow[1]);
            yearRow = Integer.parseInt(dateRow[2]);
        } catch (Exception e) {
            labelDate.setText("   Дата операции не распозналась. Отбор по дате не установлен.");
            return true;
        }

        if (selectStart.isSelected() && selectEnd.isSelected()) {

            if (yearF < yearRow) return false;
            else if (yearF == yearRow) {
                if (monthF < monthRow) return false;
                else if (monthF == monthRow) {
                    if (dayF < dayRow) return false;
                }
            }

            if (yearS < yearRow) return true;
            else if (yearS == yearRow) {
                if (monthS < monthRow) return true;
                else if (monthS == monthRow) {
                    if (dayS <= dayRow) return true;
                    else return false;
                } else return false;
            } else return false;

        } else if (!selectStart.isSelected() && selectEnd.isSelected()) {
            if (yearF > yearRow) return true;
            else if (yearF == yearRow) {
                if (monthF > monthRow) return true;
                else if (monthF == monthRow) {
                    if (dayF >= dayRow) return true;
                    else return false;
                } else return false;
            } else return false;
        } else if (selectStart.isSelected() && !selectEnd.isSelected()) {
            if (yearS < yearRow) return true;
            else if (yearS == yearRow) {
                if (monthS < monthRow) return true;
                else if (monthS == monthRow) {
                    if (dayS <= dayRow) return true;
                    else return false;
                } else return false;
            } else return false;
        } else return true;
    }

    public boolean isSelect(List<String> listINN, int i, int j, List<String> words) {
        if (listINN != null && listINN.contains(data.get(i)[j][2])) {
            if (words != null && findWords(words, data.get(i)[j][6])) {
                return true;
            } else if (words == null) {
                return true;
            }
        } else if (listINN == null) {
            if (words != null && findWords(words, data.get(i)[j][6])) {
                return true;
            } else if (words == null) {
                return true;
            }
        }
        return false;
    }

    //поиск соответствий слов пользователя в назначении
    public boolean findWords(List<String> words, String purpose) {

        for (int i = 0; i < words.size(); i++) {
            if (purpose.toLowerCase().contains(words.get(i).toLowerCase())) {
                return true;
            }
        }
        return false;
    }

    //получение списка ИНН, введенных пользователем
    public List<String> getList(JTextArea area, int num) {

        List<String> list = new ArrayList<>();

        String[] array = area.getText().split("\n");

        for (int i = 0; i < array.length; i++) {

            String INN = "";
            if (num == 0) {
                INN = array[i]
                        .replaceAll(" ", "")
                        .replaceAll("\\.", "")
                        .replaceAll(",", "");
            } else INN = array[i];

            if (INN.length() > 0) {
                list.add(INN);
            }
        }
        return list;
    }

    //получение списка ИНН по условию размера суммы по ним
    public List<String> getListINNforSum(String sumString, List<String[]> list, int INNNum, int sumNum) {

        List<String> listINN = new ArrayList<>();
        double balanceSum = 0.0;
        double INNSum = 0.0;
        double perc = 0.0;

        try {
            balanceSum = Double.parseDouble(sumString);
            sum.setBackground(Color.WHITE);
            label2Info.setText("");
        } catch (Exception e) {
            label2Info.setText(" Некорректная бал. стоимость");
            label2Info.setForeground(Color.RED);
            sum.setBackground(Color.RED);
            return listINN;
        }

        try {
            perc = Double.parseDouble(percent.getText().replaceAll(" ", ""));
            label2Info.setText("");
            percent.setBackground(Color.WHITE);
        } catch (Exception e) {
            label2Info.setText(" Некорректный процент");
            label2Info.setForeground(Color.RED);
            percent.setBackground(Color.RED);
            return listINN;
        }

        Map<String, Double> map = new HashMap<>();

        for (int i = 0; i < list.size(); i++) {
            String INN = list.get(i)[INNNum].replaceAll(" ", "");
            String INNSumString = list.get(i)[sumNum].replaceAll(" ", "")
                    .replaceAll("\\u00A0", "")
                    .replaceAll("-", ".")
                    .replaceAll(",", ".");
            try {
                if (findSum(INNSumString))
                    INNSum = Double.parseDouble(INNSumString);
                else
                    INNSum = 0.0;
                label2Info.setText("");
            } catch (Exception e) {
                labelDate.setText("  Невозможно распознать сумму в выписке");
                labelDate.setForeground(Color.RED);
                return listINN;
            }

            if (map.containsKey(INN)) map.put(INN, map.get(INN) + INNSum);
            else map.put(INN, INNSum);
        }

        for (Map.Entry pair : map.entrySet()) {
            if ((Double) pair.getValue() * 100 / balanceSum >= perc)
                listINN.add((String) pair.getKey());
        }


        return listINN;
    }

    //проверка поля на заполненность цифрами
    public boolean findSum(String sum) {
        if (sum.equals("0")) {
            return false;
        }
        Pattern pattern = Pattern.compile("\\d+");
        Matcher matcher = pattern.matcher(sum);
        while (matcher.find()) {
            return true;
        }
        return false;
    }

    //проверка, содержит ли строка дату
    public static boolean findDate(String date) {
        String[] dateRow = date.replaceAll(" ", "").split("\\.");
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("/");
        }
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("-");
        }
        if (dateRow.length < 3) {
            return false;
        }

        int dayRow = 0;
        int monthRow = 0;
        int yearRow = 0;

        try {
            dayRow = Integer.parseInt(dateRow[0]);
            monthRow = Integer.parseInt(dateRow[1]);
            yearRow = Integer.parseInt(dateRow[2]);
        } catch (Exception e) {
            return false;
        }
        return true;
    }

    //отбор строк таблицы с датой операции за определенный период (в месяцах)
    public boolean checkDate(String date, int dayB, int monthB, int yearB, int monthAgo, int yearAgo) {
        String[] dateRow = date.replaceAll(" ", "").split("\\.");
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("/");
        }
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("-");
        }
        if (dateRow.length < 3) {
            labelDate.setText("   Дата операции не распозналась (неизвестный разделитель). Отбор по дате не установлен.");
            return true;
        }

        int dayRow = 0;
        int monthRow = 0;
        int yearRow = 0;

        try {
            dayRow = Integer.parseInt(dateRow[0]);
            monthRow = Integer.parseInt(dateRow[1]);
            yearRow = dateRow[2].length() == 2 ? Integer.parseInt("20" + dateRow[2]) : Integer.parseInt(dateRow[2]);
        } catch (Exception e) {
            labelDate.setText("   Дата операции не распозналась. Отбор по дате не установлен.");
            return true;
        }


        if (yearB < yearRow + yearAgo) return true;
        else if (yearB == yearRow + yearAgo) {
            if (monthB < monthRow + monthAgo) return true;
            else if (monthB == monthRow + monthAgo) {
                if (dayB <= dayRow) return true;
                else return false;
            } else return false;
        } else if (yearB == yearRow + 1 + yearAgo) {
            if (monthRow - monthB > 12 - monthAgo) return true;
            else if (monthRow - monthB == 12 - monthAgo) {
                if (dayB <= dayRow) return true;
                else return false;
            } else return false;
        } else return false;
    }

    //отбор строк таблицы с датой в назначении платежа раньше даты подачи заявления
    public boolean checkPurpose(String purpose, int dayB, int monthB, int yearB) {
        String date = findMonth(purpose);

        if (date.equals("")) {
            date = findQuarter(purpose);
        }

        if (date.equals("")) {

            String[] pattern = {"\\d{2}\\.\\d{2}\\.\\d{4}", "\\d{2}\\.\\d{2}\\.\\d{2}", "\\d{2}\\/\\d{2}\\/\\d{4}", "\\d{2}\\/\\d{2}\\/\\d{2}"};

            List<Integer> start = new ArrayList<>();
            List<Integer> end = new ArrayList<>();

            int count = 0;

            for (int i = 0; i < 4; i++) {
                Pattern r = Pattern.compile(pattern[i]);
                Matcher m = r.matcher(purpose);

                while (m.find()) {
                    if (!start.contains(m.start())) {
                        start.add(m.start());
                        end.add(m.end());
                        count++;
                    }
                }
            }

            if (count > 1) {
                int min = 10000;
                int num = 0;
                for (int i = 0; i < start.size(); i++) {
                    num = min <= start.get(i) ? num : i;
                    min = min <= start.get(i) ? min : start.get(i);
                }


//                if (purpose.substring(end.get(num)).contains("сч") && !purpose.substring(end.get(num)).contains("дог")) {
                if (!purpose.substring(end.get(num)).contains("дог")) {
                    date = purpose.substring(start.get(num == 1 ? 0 : 1), end.get(num == 1 ? 0 : 1));
                } else {
                    date = purpose.substring(start.get(num), end.get(num));
                }
            } else if (count == 1) {
                date = purpose.substring(start.get(0), end.get(0));
            } else if (count == 0) {
                return true;
            }
        }

        String[] dateRow = date.replaceAll(" ", "").split("\\.");
        if (dateRow.length < 3) {
            dateRow = date.replace(" ", "").split("/");
        }
        if (dateRow.length < 3) {
            labelDate.setText("   Дата назначения операции не распозналась. Отбор по дате назначения не установлен.");
            return true;
        }

        int dayRow = 0;
        int monthRow = 0;
        int yearRow = 0;

        try {
            dayRow = Integer.parseInt(dateRow[0]);
            monthRow = Integer.parseInt(dateRow[1]);
            if (dateRow[2].length() == 2) {
                yearRow = Integer.parseInt("20" + dateRow[2]);
            } else {
                yearRow = Integer.parseInt(dateRow[2]);
            }
        } catch (Exception e) {
            labelDate.setText("   Дата назначения операции не распозналась. Отбор по дате назначения не установлен.");
            return true;
        }

        if (yearB > yearRow) return true;
        else if (yearB == yearRow) {
            if (monthB > monthRow) return true;
            else if (monthB == monthRow) {
                if (dayB >= dayRow) return true;
                else return false;
            } else return false;
        } else return false;
    }

    //поиск в строке месяцев
    public String findMonth(String purpose) {
        purpose = purpose.replaceAll(" ", "");

        String date = "";

        String[] pattern = {"январь\\d{4}", "февраль\\d{4}", "март\\d{4}", "апрель\\d{4}", "май\\d{4}", "июнь\\d{4}", "июль\\d{4}", "август\\d{4}", "сентябрь\\d{4}",
                "октябрь\\d{4}", "ноябрь\\d{4}", "декабрь\\d{4}"};

        for (int i = 0; i < 12; i++) {
            Pattern r = Pattern.compile(pattern[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01." + ((i + 1) > 9 ? (i + 1) : "0" + (i + 1)) + "." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        String[] patternNext = {"января\\d{4}", "февраля\\d{4}", "марта\\d{4}", "апреля\\d{4}", "мая\\d{4}", "июня\\d{4}", "июля\\d{4}", "августа\\d{4}", "сентября\\d{4}",
                "октября\\d{4}", "ноября\\d{4}", "декабря\\d{4}"};

        for (int i = 0; i < 12; i++) {
            Pattern r = Pattern.compile(patternNext[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01." + ((i + 1) > 9 ? (i + 1) : "0" + (i + 1)) + "." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        String[] patternNext1 = {"Январь\\d{4}", "Февраль\\d{4}", "Марта\\d{4}", "Апрель\\d{4}", "Май\\d{4}", "Июнь\\d{4}", "Июль\\d{4}", "Август\\d{4}", "Сентябрь\\d{4}",
                "Октябрь\\d{4}", "Ноябрь\\d{4}", "Декабрь\\d{4}"};

        for (int i = 0; i < 12; i++) {
            Pattern r = Pattern.compile(patternNext1[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01." + ((i + 1) > 9 ? (i + 1) : "0" + (i + 1)) + "." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        return date;
    }

    //поиск в строке кварталов
    public String findQuarter(String purpose) {
        purpose = purpose.replaceAll(" ", "");

        String date = "";

        String[] pattern1 = {"1кв\\d{4}", "1кв\\.\\d{4}", "1кварт\\.\\d{4}", "1кварт\\d{4}", "1квартал\\d{4}", "1к\\d{4}", "1к\\.\\d{4}",
                "Iкв\\d{4}", "Iкв\\.\\d{4}", "Iкварт\\.\\d{4}", "Iкварт\\d{4}", "Iквартал\\d{4}", "Iк\\d{4}", "Iк\\.\\d{4}"};

        for (int i = 0; i < 14; i++) {
            Pattern r = Pattern.compile(pattern1[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01.01." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        String[] pattern2 = {"2кв\\d{4}", "2кв\\.\\d{4}", "2кварт\\.\\d{4}", "2кварт\\d{4}", "2квартал\\d{4}", "2к\\d{4}", "2к\\.\\d{4}",
                "IIкв\\d{4}", "IIкв\\.\\d{4}", "IIкварт\\.\\d{4}", "IIкварт\\d{4}", "IIквартал\\d{4}", "IIк\\d{4}", "IIк\\.\\d{4}"};

        for (int i = 0; i < 14; i++) {
            Pattern r = Pattern.compile(pattern2[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01.04." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        String[] pattern3 = {"3кв\\d{4}", "3кв\\.\\d{4}", "3кварт\\.\\d{4}", "3кварт\\d{4}", "3квартал\\d{4}", "3к\\d{4}", "3к\\.\\d{4}",
                "IIIкв\\d{4}", "IIIкв\\.\\d{4}", "IIIкварт\\.\\d{4}", "IIIкварт\\d{4}", "IIIквартал\\d{4}", "IIIк\\d{4}", "IIIк\\.\\d{4}"};

        for (int i = 0; i < 14; i++) {
            Pattern r = Pattern.compile(pattern3[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01.07." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }

        String[] pattern4 = {"4кв\\d{4}", "4кв\\.\\d{4}", "4кварт\\.\\d{4}", "4кварт\\d{4}", "4квартал\\d{4}", "4к\\d{4}", "4к\\.\\d{4}",
                "IVкв\\d{4}", "IVкв\\.\\d{4}", "IVкварт\\.\\d{4}", "IVкварт\\d{4}", "IVквартал\\d{4}", "IVк\\d{4}", "IVк\\.\\d{4}"};

        for (int i = 0; i < 14; i++) {
            Pattern r = Pattern.compile(pattern4[i]);
            Matcher m = r.matcher(purpose);

            if (m.find()) {
                date = "01.10." + purpose.substring(m.start(), m.end()).substring(purpose.substring(m.start(), m.end()).length() - 4);
                return date;
            }
        }
        return date;
    }

    //конструктор формы
    public Main() {

        //создаем каркас: основная форма, вкладки
        super("Анализ сделок");
        this.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);

        setBounds(100, 100, 1200, 700);

        this.getContentPane().setLayout(new BoxLayout(this.getContentPane(), BoxLayout.Y_AXIS));

        tabbedPaneFirst.addTab("Получение данных", jpanel);
        tabbedPaneFirst.addTab("Анализ данных", jpanelRez);
        tabbedPaneFirst.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        jpanelRez.setLayout(new BoxLayout(jpanelRez, BoxLayout.Y_AXIS));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа от верхней границы формы
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(Box.createHorizontalGlue());
        this.add(panel);
        this.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа от верхней границы формы
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(Box.createHorizontalGlue());
        jpanelRez.add(panel);
        jpanelRez.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода даты и всех отборов в анализе (вверху второй вкладки)
        JPanel panelSelect = new JPanel();
        panelSelect.setLayout(new GridLayout(1, 4));
        panelSelect.setMaximumSize(new Dimension(3000, 100));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода даты подачи заявления о банкротстве
        //JPanel panelDate = new JPanel();
        panelDate.setLayout(new GridLayout(3, 2));
        panelDate.setBorder(BorderFactory.createTitledBorder("Дата"));
        ((javax.swing.border.TitledBorder) panelDate.getBorder()).
                setTitleFont(new Font("Arial", Font.BOLD, 14));
        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        JLabel label = new JLabel(" Число");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        day = new JTextField();
        day.setColumns(4);
        day.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panel.add(day, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Месяц (числом)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        month = new JTextField();
        month.setColumns(4);
        month.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(month, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Год (полностью)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        year = new JTextField();
        year.setColumns(4);
        year.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(year, BorderLayout.WEST);
        panelDate.add(panel);

        panelSelect.add(panelDate);
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода списка ИНН
        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        selectINN.setText("Отбор по ИНН");
        selectINN.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(selectINN, BorderLayout.NORTH);
        areaINN.setRows(3);
        areaINN.setColumns(30);
        areaINN.setBorder(BorderFactory.createTitledBorder("ИНН"));
        areaINN.setToolTipText("ИНН вводить в столбик");
        areaINN.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        ((javax.swing.border.TitledBorder) areaINN.getBorder()).
                setTitleFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        JScrollPane pane = new JScrollPane(areaINN,
                JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        panel.add(pane);
        panelSelect.add(panel);
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода ключевых слов в назначении
        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        selectPurpose.setText("Отбор по назначению");
        selectPurpose.setFont(new Font("Arial", Font.BOLD, 14));

        panel.add(selectPurpose, BorderLayout.NORTH);
        areaPurpose.setRows(3);
        areaPurpose.setColumns(30);
        areaPurpose.setBorder(BorderFactory.createTitledBorder("Назначение"));
        areaPurpose.setToolTipText("Ключевые слова в назначении платежа, по которым необходим поиск, вводить в столбик");
        areaPurpose.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        ((javax.swing.border.TitledBorder) areaPurpose.getBorder()).
                setTitleFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        pane = new JScrollPane(areaPurpose,
                JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        panel.add(pane);
        panelSelect.add(panel);
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода стоимости балансовой
        JPanel panelSum = new JPanel();
        panelSum.setBorder(BorderFactory.createEtchedBorder());
        panelSum.setLayout(new BoxLayout(panelSum, BoxLayout.Y_AXIS));

        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        selectIBalanceSum.setText("Отбор по балансовой стоимости");

        selectIBalanceSum.addActionListener(new SetWhite(list1, label2Info));

        selectIBalanceSum.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(selectIBalanceSum);
        panel.add(Box.createHorizontalGlue());
        panelSum.add(panel);
        panelSum.add(Box.createRigidArea(new Dimension(0, 10)));

        JPanel panelTab = new JPanel();
        panelTab.setLayout(new GridLayout(2, 1));

        panel = new JPanel();
        label = new JLabel("  Бал. ст-ость  ");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label);
        //sum = new JTextField();
        sum.setColumns(7);
        sum.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(sum);
        panelTab.add(panel);

        panel = new JPanel();
        label = new JLabel("  % от суммы по ИНН не менее ");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label);
        //percent = new JTextField();
        percent.setColumns(2);
        percent.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        percent.setText("10");
        panel.add(percent);
        panelTab.add(panel);

        panelSum.add(panelTab);
        panelSum.add(Box.createRigidArea(new Dimension(0, 10)));

        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        label2Info.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(label2Info);
        panel.add(Box.createHorizontalGlue());
        panelSum.add(panel);
        panelSum.add(Box.createRigidArea(new Dimension(0, 10)));
        panelSelect.add(panelSum);
///////////////////////////////////////////////////////////////////////////////////////////////////
        jpanelRez.add(panelSelect);
        jpanelRez.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        //информация о корректности ввода даты
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        labelDate.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(labelDate);
        panel.add(Box.createHorizontalGlue());
        jpanelRez.add(panel);
        jpanelRez.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        //вкладки для разных видов анализа
        JPanel panel1 = new JPanel();
        panel1.setLayout(new BoxLayout(panel1, BoxLayout.Y_AXIS));
        JPanel panel2 = new JPanel();
        panel2.setLayout(new BoxLayout(panel2, BoxLayout.Y_AXIS));

        tabbedPaneResult.addTab("Сделки с предпочтением", panel1);
        tabbedPaneResult.addTab("Сделки с неравноценным встречным исполнением", panel2);
        tabbedPaneResult.addTab("Сделки с злоупотреблением правом", jpanelRez5);
        tabbedPaneResult.addTab("Универсальный анализ сделок", jpanelRez6);
        tabbedPaneResult.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        tabbedPaneResult.addMouseListener(new Block());//.addFocusListener(new Block());

        jpanelRez.add(tabbedPaneResult);
        jpanelRez.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData.addActionListener(new SaveData(0, "полный анализ", labelSaveData));
        saveData.setFont(new Font("Arial", Font.BOLD, 14));
        labelSaveData.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(saveData);
        panel.add(labelSaveData);
        panel.add(Box.createHorizontalGlue());
        jpanelRez.add(panel);
        jpanelRez.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        //вкладки для сделок с предпочтением
        tabbedPane1.addTab("за месяц", jpanelRez1);
        tabbedPane1.addTab("за полгода", jpanelRez2);
        tabbedPane1.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panel1.add(tabbedPane1);
        panel1.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        //вкладки для сделок с неравноценным встречным исполнением
        tabbedPane2.addTab("за год", jpanelRez3);
        tabbedPane2.addTab("за 3 года", jpanelRez4);
        tabbedPane2.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panel2.add(tabbedPane2);
        panel2.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        jpanel.setLayout(new BoxLayout(jpanel, BoxLayout.Y_AXIS));
        jpanelRez1.setLayout(new BoxLayout(jpanelRez1, BoxLayout.Y_AXIS));
        jpanelRez2.setLayout(new BoxLayout(jpanelRez2, BoxLayout.Y_AXIS));
        jpanelRez3.setLayout(new BoxLayout(jpanelRez3, BoxLayout.Y_AXIS));
        jpanelRez4.setLayout(new BoxLayout(jpanelRez4, BoxLayout.Y_AXIS));
        jpanelRez5.setLayout(new BoxLayout(jpanelRez5, BoxLayout.Y_AXIS));
        jpanelRez6.setLayout(new BoxLayout(jpanelRez6, BoxLayout.Y_AXIS));
///////////////////////////////////////////////////////////////////////////////////////////////

// пустая панель для отступа от верхней границы формы
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez1.add(panel);
        jpanelRez1.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделак за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search1.addActionListener(new search(1));
        search1.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search1);
        label1.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(label1);
        panel.add(Box.createHorizontalGlue());
        jpanelRez1.add(panel);
        jpanelRez1.add(Box.createRigidArea(new Dimension(0, 10)));
////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData1.addActionListener(new SaveData(1, "за месяц", labelSave1));
        saveData1.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData1);
        labelSave1.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave1);
        panel.add(Box.createHorizontalGlue());
        jpanelRez1.add(panel);
        jpanelRez1.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок за месяц
        panelSearch1.setLayout(new BorderLayout());
        JTable jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane1 = new JScrollPane(jTable);
        panelSearch1.add(scrollPane1);
        jpanelRez1.add(panelSearch1);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
//        panel = new JPanel();
//        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
//        saveData1.addActionListener(new SaveData(1, "за месяц", labelSave1));
//        saveData1.setFont(new Font("Arial", Font.BOLD, 14));
//        panel.add(saveData1);
//        labelSave1.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
//        panel.add(labelSave1);
//        panel.add(Box.createHorizontalGlue());
//        jpanelRez1.add(panel);
//        jpanelRez1.add(Box.createRigidArea(new Dimension(0, 10)));
//////////////////////////////////////////////////////////////////////////////////////////////////

        // пустая панель для отступа
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez2.add(panel);
        jpanelRez2.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделак за полгода
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search2.addActionListener(new search(2));
        search2.setFont(new Font("Arial", Font.BOLD, 14));
        label2.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search2);
        panel.add(label2);
        panel.add(Box.createHorizontalGlue());
        jpanelRez2.add(panel);
        jpanelRez2.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData2.addActionListener(new SaveData(2, "за полгода", labelSave2));
        saveData2.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData2);
        labelSave2.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave2);
        panel.add(Box.createHorizontalGlue());
        jpanelRez2.add(panel);
        jpanelRez2.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок за полгода
        panelSearch2.setLayout(new BorderLayout());
        jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane2 = new JScrollPane(jTable);
        panelSearch2.add(scrollPane2);
        jpanelRez2.add(panelSearch2);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez3.add(panel);
        jpanelRez3.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделак за год
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search3.addActionListener(new search(3));
        search3.setFont(new Font("Arial", Font.BOLD, 14));
        label3.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search3);
        panel.add(label3);
        panel.add(Box.createHorizontalGlue());
        jpanelRez3.add(panel);
        jpanelRez3.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData3.addActionListener(new SaveData(3, "за год", labelSave3));
        saveData3.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData3);
        labelSave3.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave3);
        panel.add(Box.createHorizontalGlue());
        jpanelRez3.add(panel);
        jpanelRez3.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок за год
        panelSearch3.setLayout(new BorderLayout());
        jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane3 = new JScrollPane(jTable);
        panelSearch3.add(scrollPane3);
        jpanelRez3.add(panelSearch3);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez4.add(panel);
        jpanelRez4.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделак за 3 года
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search4.addActionListener(new search(4));
        search4.setFont(new Font("Arial", Font.BOLD, 14));
        label4.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search4);
        panel.add(label4);
        panel.add(Box.createHorizontalGlue());
        jpanelRez4.add(panel);
        jpanelRez4.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData4.addActionListener(new SaveData(4, "за 3 года", labelSave4));
        saveData4.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData4);
        labelSave4.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave4);
        panel.add(Box.createHorizontalGlue());
        jpanelRez4.add(panel);
        jpanelRez4.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок за 3 года
        panelSearch4.setLayout(new BorderLayout());
        jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane4 = new JScrollPane(jTable);
        panelSearch4.add(scrollPane4);
        jpanelRez4.add(panelSearch4);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez5.add(panel);
        jpanelRez5.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделак за 10 лет
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search5.addActionListener(new search(5));
        search5.setFont(new Font("Arial", Font.BOLD, 14));
        label5.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search5);
        panel.add(label5);
        panel.add(Box.createHorizontalGlue());
        jpanelRez5.add(panel);
        jpanelRez5.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData5.addActionListener(new SaveData(5, "за 10 лет", labelSave5));
        saveData5.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData5);
        labelSave5.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave5);
        panel.add(Box.createHorizontalGlue());
        jpanelRez5.add(panel);
        jpanelRez5.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок за 10 лет
        panelSearch5.setLayout(new BorderLayout());
        jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane5 = new JScrollPane(jTable);
        panelSearch5.add(scrollPane5);
        jpanelRez5.add(panelSearch5);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода даты и всех отборов в анализе (вверху второй вкладки)
        panelSelect = new JPanel();
        panelSelect.setLayout(new GridLayout(1, 4));
        panelSelect.setMaximumSize(new Dimension(3000, 100));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода начала периода
        JPanel panelStart = new JPanel();
        panelStart.setLayout(new BorderLayout());
        panelStart.setBorder(BorderFactory.createEtchedBorder());
        selectStart.setText("Установить начало периода");
        selectStart.setFont(new Font("Arial", Font.BOLD, 14));


        selectStart.addActionListener(new SetWhite(list2, label6));

        panelStart.add(selectStart, BorderLayout.NORTH);

        panelDate = new JPanel();
        panelDate.setLayout(new GridLayout(3, 2));

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Число");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //dayStart = new JTextField();
        dayStart.setColumns(4);
        dayStart.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panel.add(dayStart, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Месяц (числом)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //monthStart = new JTextField();
        monthStart.setColumns(4);
        monthStart.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(monthStart, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Год (полностью)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //yearStart = new JTextField();
        yearStart.setColumns(4);
        yearStart.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(yearStart, BorderLayout.WEST);
        panelDate.add(panel);

        panelStart.add(panelDate);
        panelSelect.add(panelStart);

        jpanelRez6.add(panelSelect);
        jpanelRez6.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода конца периода
        panelStart = new JPanel();
        panelStart.setLayout(new BorderLayout());
        panelStart.setBorder(BorderFactory.createEtchedBorder());
        selectEnd.setText("Установить окончание периода");
        selectEnd.setFont(new Font("Arial", Font.BOLD, 14));


        selectEnd.addActionListener(new SetWhite(list3, label6));

        panelStart.add(selectEnd, BorderLayout.NORTH);

        panelDate = new JPanel();
        panelDate.setLayout(new GridLayout(3, 2));

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Число");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //dayEnd = new JTextField();
        dayEnd.setColumns(4);
        dayEnd.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panel.add(dayEnd, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Месяц (числом)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //monthEnd = new JTextField();
        monthEnd.setColumns(4);
        monthEnd.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(monthEnd, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        label = new JLabel(" Год (полностью)");
        label.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(label, BorderLayout.WEST);
        panelDate.add(panel);

        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        //yearEnd = new JTextField();
        yearEnd.setColumns(4);
        yearEnd.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(yearEnd, BorderLayout.WEST);
        panelDate.add(panel);

        panelStart.add(panelDate);
        panelSelect.add(panelStart);
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для ввода ФИО
        panel = new JPanel();
        panel.setLayout(new BorderLayout());
        selectFIO.setText("Отбор по наименованию/Ф.И.О.");
        selectFIO.setFont(new Font("Arial", Font.BOLD, 14));

        panel.add(selectFIO, BorderLayout.NORTH);
        areaFIO.setRows(4);
        areaFIO.setColumns(30);
        areaFIO.setToolTipText("Наименование или ФИО плательщика/получателя денежных средств, по которым необходим поиск, вводить в столбик");
        areaFIO.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        pane = new JScrollPane(areaFIO,
                JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        panel.add(pane);
        panelSelect.add(panel);
        panelSelect.add(Box.createHorizontalGlue());

        jpanelRez6.add(panelSelect);
        jpanelRez6.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для поиска сделок универсальный отбор
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        search6.addActionListener(new searchAll());
        search6.setFont(new Font("Arial", Font.BOLD, 14));
        label6.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(search6);
        panel.add(label6);
        panel.add(Box.createHorizontalGlue());
        jpanelRez6.add(panel);
        jpanelRez6.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выгрузки сделок за месяц
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        saveData6.addActionListener(new SaveData(6, "универсальный", labelSave6));
        saveData6.setFont(new Font("Arial", Font.BOLD, 14));
        panel.add(saveData6);
        labelSave6.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.add(labelSave6);
        panel.add(Box.createHorizontalGlue());
        jpanelRez6.add(panel);
        jpanelRez6.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для вывода найденных сделок универсальный отбор
        panelSearch6.setLayout(new BorderLayout());
        jTable = new JTable(tableData, headers);
        jTable.getTableHeader().setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        tablesRez.add(jTable);
        scrollPane6 = new JScrollPane(jTable);
        panelSearch6.add(scrollPane6);
        jpanelRez6.add(panelSearch6);
/////////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа от верхней границы формы
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(Box.createHorizontalGlue());
        jpanel.add(panel);
        jpanel.add(Box.createRigidArea(new Dimension(0, 10)));
/////////////////////////////////////////////////////////////////////////////////////////////////
        // панель для выбора файла
        panel = new JPanel();
        open.addActionListener(new OpenL());
        open.setFont(new Font("Arial", Font.BOLD, 14));
        fileName.setEnabled(false);
        fileName.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel("   "));
        panel.add(open);
        panel.add(new JLabel("   "));
        panel.add(fileName);
        panel.add(Box.createHorizontalGlue());
        jpanel.add(panel);
        jpanel.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // открытие файла
        panel = new JPanel();
        openExcel.addActionListener(new OpenExcel());
        openExcel.setFont(new Font("Arial", Font.BOLD, 14));
        openExcel.setEnabled(false);
        labelOpenExcel.setEnabled(false);
        labelOpenExcel.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel("   "));
        panel.add(openExcel);
        panel.add(new JLabel("   "));
        panel.add(labelOpenExcel);
        panel.add(Box.createHorizontalGlue());
        jpanel.add(panel);
        jpanel.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // создание первой вкладки (для первой страницы)
        jpanel.add(tabbedPane);
        JPanel panelTabbedPane = new JPanel();
        tabbedPane.addTab("Лист 1", panelTabbedPane);
        tabbedPane.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));

        panelTabbedPane.setLayout(new BoxLayout(panelTabbedPane, BoxLayout.Y_AXIS));
///////////////////////////////////////////////////////////////////////////////////////////////////
        // пустая панель для отступа от верхней границы формы на второй вкладке с результатом
        panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
        panel.add(new JLabel(""));
        panel.add(Box.createHorizontalGlue());
        jpanelRez1.add(panel);
        jpanelRez1.add(Box.createRigidArea(new Dimension(0, 10)));
///////////////////////////////////////////////////////////////////////////////////////////////////

        this.add(tabbedPaneFirst);
        this.setVisible(true);
    }

private class OpenExcel implements ActionListener {
    public void actionPerformed(ActionEvent e) {
        Runtime r = Runtime.getRuntime();
        try {
            Desktop.getDesktop().open(new File(fileName.getText()));
        } catch (Exception exc) {
            labelOpenExcel.setText("Проблемка с открытием файла...");
        }
    }
}

private class SaveData implements ActionListener {
    public int number;
    public JLabel label;
    public String nameFile;

    public SaveData(int number, String nameFile, JLabel label) {
        this.number = number;
        this.label = label;
        this.nameFile = nameFile;
    }

    public void actionPerformed(ActionEvent e) {
        XSSFWorkbook workbook = new XSSFWorkbook();

        JFileChooser c = new JFileChooser();
        if (directory.contains("\\")) {
            c = new JFileChooser(directory);
        }
        c.setDialogTitle("Выбор каталога для сохранения результата анализа");
        c.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = c.showOpenDialog(Main.this);

        if (result == JFileChooser.APPROVE_OPTION) {
            path = c.getSelectedFile().getAbsolutePath() + "\\" + name + "_" + nameFile + "_v" + count + ".xlsx";
            label.setText(path);
            count++;
        }
        if (result == JFileChooser.CANCEL_OPTION) {
            label.setText("   Файл не выбран...");
            return;
        }
        if (result == JFileChooser.ERROR_OPTION) {
            label.setText("   Ошибочка при выборе каталога(((");
            return;
        }
        if (number > 0) {

            XSSFSheet sheet = workbook.createSheet(text[number - 1]);

            XSSFFont sheetTitleFont = workbook.createFont();
            XSSFCellStyle cellTitleStyle = workbook.createCellStyle();
            sheetTitleFont.setBold(true);
            cellTitleStyle.setFont(sheetTitleFont);

            TableModel model = tablesRez.get(number - 1).getModel();

            //Get Header
            TableColumnModel tcm = tablesRez.get(number - 1).getColumnModel();
            XSSFRow row = sheet.createRow(0);
            for (int j = 0; j < tcm.getColumnCount(); j++) {
                XSSFCell cell = row.createCell(j);
                cell.setCellValue(tcm.getColumn(j).getHeaderValue().toString());
                cell.setCellStyle(cellTitleStyle);

            }

            //Get Other details
            for (int i = 0; i < model.getRowCount(); i++) {
                XSSFRow fRow = sheet.createRow(i + 1);
                for (int j = 0; j < model.getColumnCount(); j++) {
                    XSSFCell cell = fRow.createCell(j);
                    cell.setCellValue(model.getValueAt(i, j).toString());
                }
            }
        } else {

            for (int i = 0; i < 6; i++) {
                XSSFSheet sheet = workbook.createSheet(text[i]);

                XSSFFont sheetTitleFont = workbook.createFont();
                XSSFCellStyle cellTitleStyle = workbook.createCellStyle();
                sheetTitleFont.setBold(true);
                cellTitleStyle.setFont(sheetTitleFont);

                TableModel model = tablesRez.get(i).getModel();

                //Get Header
                TableColumnModel tcm = tablesRez.get(i).getColumnModel();
                XSSFRow row = sheet.createRow(0);
                for (int j = 0; j < tcm.getColumnCount(); j++) {
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue(tcm.getColumn(j).getHeaderValue().toString());
                    cell.setCellStyle(cellTitleStyle);

                }

                //Get Other details
                for (int k = 0; k < model.getRowCount(); k++) {
                    XSSFRow fRow = sheet.createRow(k + 1);
                    for (int j = 0; j < model.getColumnCount(); j++) {
                        XSSFCell cell = fRow.createCell(j);
                        cell.setCellValue(model.getValueAt(k, j).toString());
                    }
                }
            }
        }

        File xlsx = new File(path);

        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(xlsx)) {
            workbook.write(out);
            label.setText("   Файл успешно сохранен по адресу: " + path);
        } catch (IOException ex) {
            label.setText(label.getText() + ". Ошибочка при сохранении файла. Возможно файл с таким именем уже создан в выбранном каталоге.");
        }
    }

}

    public static void main(String[] args) throws Exception {
        //JDialog d = new JDialog();
        //d.setTitle("Предупреждение");
        //d.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        //d.setLayout(new GridLayout(2,1));
        //JLabel l = new JLabel("  Тестовый период работы программы закончился.");
        //l.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        //d.add(l);
        //l = new JLabel("  Обратитесь к разработчику.");
        //l.setFont(new Font("Arial", Font.TRUETYPE_FONT, 14));
        //d.add(l);

        //LocalDateTime trialPeriod = LocalDateTime.of(2020,03,25,18,00);

        //if (trialPeriod.isBefore(LocalDateTime.now())) {
        //    d.setBounds(500, 300, 400, 100);
        //    d.setVisible(true);
        //    return;
        //}

        nf.setMaximumFractionDigits(2);
        nf.setGroupingUsed(false);

        list1.add(sum);
        list1.add(percent);

        list2.add(dayStart);
        list2.add(monthStart);
        list2.add(yearStart);

        list3.add(dayEnd);
        list3.add(monthEnd);
        list3.add(yearEnd);

        text[0] = "Сделки с предпочт. за месяц";
        text[1] = "Сделки с предпочт.за полгода";
        text[2] = "Неравн. встречн. исп. за год";
        text[3] = "Неравн. встречн. исп. за 3 года";
        text[4] = "Злоупотр. правом за 10 лет";
        text[5] = "Универсальный анализ сделок";

        try {
            UIManager.LookAndFeelInfo[] available =
                    UIManager.getInstalledLookAndFeels();
            List<String> names = new ArrayList<>();
            for (UIManager.LookAndFeelInfo info : available) {
                names.add(info.getName());
                if (info.getName().equals("Nimbus")) {
                    UIManager.setLookAndFeel(info.getClassName());
                }
            }


        } catch (Exception e) {
            e.printStackTrace();
        }

        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                prog = new Main();
            }
        });

        MapInit();

    }

    //функция преобразования названия колонки Excel в число
    private static void MapInit() {
        map.put("A", 1);
        map.put("B", 2);
        map.put("C", 3);
        map.put("D", 4);
        map.put("E", 5);
        map.put("F", 6);
        map.put("G", 7);
        map.put("H", 8);
        map.put("I", 9);
        map.put("J", 10);
        map.put("K", 11);
        map.put("L", 12);
        map.put("M", 13);
        map.put("N", 14);
        map.put("O", 15);
        map.put("P", 16);
        map.put("Q", 17);
        map.put("R", 18);
        map.put("S", 19);
        map.put("T", 20);
        map.put("U", 21);
        map.put("V", 22);
        map.put("W", 23);
        map.put("X", 24);
        map.put("Y", 25);
        map.put("Z", 26);
        map.put("AA", 27);
        map.put("AB", 28);
        map.put("AC", 29);
        map.put("AD", 30);
        map.put("AE", 31);
        map.put("AF", 32);
        map.put("AG", 33);
        map.put("AH", 34);
        map.put("AI", 35);
        map.put("AJ", 36);
        map.put("AK", 37);
        map.put("AL", 38);
        map.put("AM", 39);
        map.put("AN", 40);
        map.put("AO", 41);
        map.put("AP", 42);
        map.put("AQ", 43);
        map.put("AR", 44);
        map.put("AS", 45);
        map.put("AT", 46);
        map.put("AU", 47);
        map.put("AV", 48);
        map.put("AW", 49);
        map.put("AX", 50);
        map.put("AY", 51);
        map.put("AZ", 52);
        map.put("BA", 53);
        map.put("BB", 54);
        map.put("BC", 55);
        map.put("BD", 56);
        map.put("BE", 57);
        map.put("BF", 58);
        map.put("BG", 59);
        map.put("BH", 60);
        map.put("BI", 61);
        map.put("BJ", 62);
        map.put("BK", 63);
        map.put("BL", 64);
        map.put("BM", 65);
        map.put("BN", 66);
        map.put("BO", 67);
        map.put("BP", 68);
        map.put("BQ", 69);
        map.put("BR", 70);
        map.put("BS", 71);
        map.put("BT", 72);
        map.put("BU", 73);
        map.put("BV", 74);
        map.put("BW", 75);
        map.put("BX", 76);
        map.put("BY", 77);
        map.put("BZ", 78);
    }
}
