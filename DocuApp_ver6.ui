<?xml version="1.0" encoding="UTF-8"?>
<ui version="4.0">
 <class>MainWindow</class>
 <widget class="QMainWindow" name="MainWindow">
  <property name="geometry">
   <rect>
    <x>0</x>
    <y>0</y>
    <width>1000</width>
    <height>650</height>
   </rect>
  </property>
  <property name="windowTitle">
   <string>TARDIS</string>
  </property>
  <property name="styleSheet">
   <string notr="true">
    QMainWindow { background: transparent; }
    QWidget#centralwidget { background: #16161a; border: 1px solid #3a3a40; border-radius: 12px; }
    QFrame#headerr, QFrame#footer { background: transparent; border: none; }
    QFrame#selectionframe, QFrame#generate_frame { background: #202025; border: 1px solid #3a3a40; border-radius: 8px; }
    
    #centralwidget QPushButton { padding: 0px 15px; background: #0085ca; color: white; border-radius: 5px; font-weight: bold; min-height: 28px; }
    #centralwidget QPushButton:hover { background: #3c649f; }
    #centralwidget QPushButton:disabled { background: #333333; color: #777777; }
    
    #headerr QPushButton { padding: 0px; }
    #crop_button { padding: 0px; }
    
    QPushButton#btn_tab_perf, QPushButton#btn_tab_wave {
        padding: 0px; background-color: transparent; color: #777777; border: none; border-bottom: 3px solid transparent; border-radius: 0px; font-size: 14px;
    }
    QPushButton#btn_tab_perf:hover, QPushButton#btn_tab_wave:hover { background-color: #2a2a30; }
    QPushButton#btn_tab_perf:checked, QPushButton#btn_tab_wave:checked { color: white; border-bottom: 3px solid #0085ca; }

    QLineEdit { background: #1a1a1e; color: #e0e0e0; border-radius: 4px; padding-left: 5px; border: 1px solid #3a3a40; }
    QLineEdit:disabled { background: #2a2a2e; color: #777777; }
    
    /* --- TREE VIEW AND LIST IMPROVEMENTS --- */
    #centralwidget QListWidget, #centralwidget QTreeWidget { 
        background: #121214; 
        color: #f0f0f0; 
        border: 1px solid #3a3a40; 
        border-radius: 5px; 
        font-size: 13px; 
    }
    #centralwidget QTreeWidget::item { padding: 4px 8px; min-height: 20px; }
    #centralwidget QTreeWidget::item:selected { background-color: #0085ca; color: white; }
    #centralwidget QTreeWidget::item:hover:!selected { background-color: #2a2a30; }

    #centralwidget QLabel { color: white; border: none; }
    QLabel#data_exporter { padding-left: 6px; padding-top: 12px; }

    /* --- EXPANDED DIALOG AND MESSAGEBOX STYLES --- */
    QDialog#expandedPreviewDialog { background-color: #16161a; color: white; }
    QLabel#expandedTitleLabel { color: #0085ca; font-size: 14pt; font-weight: bold; margin-bottom: 10px; }
    QLabel#expandedImageLabel { background: #0b0b0d; border: 2px solid #3a3a40; border-radius: 8px; }
    QPushButton#expandedCropButton { background: #0085ca; color: white; border-radius: 4px; padding: 8px 25px; font-weight: bold; }
    QPushButton#expandedCropButton:disabled { background: #333333; color: #777777; }
    
    QMessageBox { background-color: #16161a; }
    QMessageBox QLabel { color: white; font-size: 13px; font-weight: bold; }
    QMessageBox QPushButton { background-color: #0085ca; color: white; border-radius: 4px; padding: 6px 20px; font-weight: bold; min-width: 60px; }
    QMessageBox QPushButton:hover { background-color: #3c649f; }
   </string>
  </property>
  <widget class="QWidget" name="centralwidget">
   <layout class="QVBoxLayout" name="verticalLayout_main">
    <property name="spacing"><number>0</number></property>
    <property name="margin"><number>0</number></property>
    <item>
     <widget class="QFrame" name="headerr">
      <property name="minimumSize"><size><width>0</width><height>50</height></size></property>
      <layout class="QHBoxLayout" name="horizontalLayout_header">
       <item><widget class="QLabel" name="data_exporter"><property name="font"><font><family>Bahnschrift</family><pointsize>18</pointsize><bold>true</bold></font></property><property name="text"><string>T.A.R.D.I.S.</string></property></widget></item>
       <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
       <item><widget class="QPushButton" name="minimize_button"><property name="maximumSize"><size><width>30</width><height>30</height></size></property><property name="text"><string>—</string></property></widget></item>
       <item><widget class="QPushButton" name="maximize_button"><property name="maximumSize"><size><width>30</width><height>30</height></size></property><property name="text"><string>▢</string></property></widget></item>
       <item><widget class="QPushButton" name="exit_button"><property name="maximumSize"><size><width>30</width><height>30</height></size></property><property name="text"><string>✕</string></property></widget></item>
      </layout>
     </widget>
    </item>
    <item>
     <layout class="QHBoxLayout" name="horizontalLayout_body">
      <property name="margin"><number>15</number></property>
      <property name="spacing"><number>15</number></property>
      
      <item>
       <widget class="QFrame" name="selectionframe">
        <layout class="QVBoxLayout" name="verticalLayout_selection">
         <property name="margin"><number>15</number></property>
         <property name="spacing"><number>15</number></property>
         
         <item>
          <layout class="QHBoxLayout" name="tab_layout">
           <item>
            <widget class="QPushButton" name="btn_tab_perf">
             <property name="minimumSize"><size><width>0</width><height>40</height></size></property>
             <property name="text"><string>PERFORMANCE DATA</string></property>
             <property name="checkable"><bool>true</bool></property>
             <property name="checked"><bool>true</bool></property>
            </widget>
           </item>
           <item>
            <widget class="QPushButton" name="btn_tab_wave">
             <property name="minimumSize"><size><width>0</width><height>40</height></size></property>
             <property name="text"><string>WAVEFORMS</string></property>
             <property name="checkable"><bool>true</bool></property>
            </widget>
           </item>
          </layout>
         </item>
         
         <item>
          <widget class="QStackedWidget" name="stackedWidget">
           <property name="currentIndex"><number>0</number></property>
           
           <!-- PAGE 0: PERFORMANCE DATA -->
           <widget class="QWidget" name="page_perf">
            <layout class="QVBoxLayout" name="layout_page_perf">
             <property name="margin"><number>0</number></property>
             <item>
              <layout class="QHBoxLayout">
               <item><widget class="QPushButton" name="performancedata_sel">
                <property name="minimumSize"><size><width>90</width><height>0</height></size></property>
                <property name="text"><string>Add File</string></property>
               </widget></item>
               <item><widget class="QLineEdit" name="performancedata_path"></widget></item>
              </layout>
             </item>
             <item>
              <widget class="QTreeWidget" name="performance_tree">
               <property name="sizePolicy"><sizepolicy hsizetype="Expanding" vsizetype="Expanding"><horstretch>1</horstretch><verstretch>1</verstretch></sizepolicy></property>
               <property name="headerHidden"><bool>true</bool></property>
               <column><property name="text"><string>Name</string></property></column>
              </widget>
             </item>
             <item>
              <layout class="QHBoxLayout">
               <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
               <item><widget class="QPushButton" name="refresh_button_perf">
                <property name="minimumSize"><size><width>120</width><height>30</height></size></property>
                <property name="text"><string>REFRESH LIST</string></property>
               </widget></item>
               <item><widget class="QPushButton" name="clear_perf_button">
                <property name="minimumSize"><size><width>120</width><height>30</height></size></property>
                <property name="text"><string>CLEAR DATA</string></property>
               </widget></item>
               <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
              </layout>
             </item>
            </layout>
           </widget>
           
           <!-- PAGE 1: WAVEFORMS -->
           <widget class="QWidget" name="page_wave">
            <layout class="QVBoxLayout" name="layout_page_wave">
             <property name="margin"><number>0</number></property>
             <item>
              <layout class="QHBoxLayout">
               <item><widget class="QPushButton" name="waveforms_add_folder">
                <property name="minimumSize"><size><width>90</width><height>0</height></size></property>
                <property name="text"><string>Add File</string></property>
               </widget></item>
               <item><widget class="QLineEdit" name="waveforms_path"></widget></item>
              </layout>
             </item>
             <item>
              <widget class="QTreeWidget" name="waveform_tree">
               <property name="sizePolicy"><sizepolicy hsizetype="Expanding" vsizetype="Expanding"><horstretch>1</horstretch><verstretch>1</verstretch></sizepolicy></property>
               <property name="headerHidden"><bool>true</bool></property>
               <column><property name="text"><string>Name</string></property></column>
              </widget>
             </item>
             <item>
              <layout class="QHBoxLayout">
               <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
               <item><widget class="QPushButton" name="refresh_button_wave">
                <property name="minimumSize"><size><width>120</width><height>30</height></size></property>
                <property name="text"><string>REFRESH LIST</string></property>
               </widget></item>
               <item><widget class="QPushButton" name="waveforms_clear_folders">
                <property name="minimumSize"><size><width>120</width><height>30</height></size></property>
                <property name="text"><string>CLEAR WAVEFORMS</string></property>
               </widget></item>
               <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
              </layout>
             </item>
            </layout>
           </widget>
           
          </widget>
         </item>
        </layout>
       </widget>
      </item>
      
      <!-- RIGHT PANEL (Fixed Preview) -->
      <item>
       <widget class="QFrame" name="generate_frame">
        <property name="maximumSize"><size><width>400</width><height>16777215</height></size></property>
        <layout class="QVBoxLayout" name="verticalLayout_gen">
         <property name="margin"><number>15</number></property>
         <item><widget class="QLabel" name="images_preview_text">
          <property name="font"><font><bold>true</bold><pointsize>11</pointsize></font></property>
          <property name="text"><string>PREVIEW</string></property>
          <property name="alignment"><set>Qt::AlignCenter</set></property>
         </widget></item>
         <item><widget class="QLabel" name="file_view">
          <property name="sizePolicy"><sizepolicy hsizetype="Expanding" vsizetype="Expanding"><horstretch>0</horstretch><verstretch>1</verstretch></sizepolicy></property>
          <property name="minimumSize"><size><width>0</width><height>150</height></size></property>
          <property name="styleSheet"><string>background: #0b0b0d; border-radius: 4px; border: 1px solid #3a3a40;</string></property>
         </widget></item>
         <item>
          <layout class="QGridLayout" name="gridLayout_crop">
           <item row="0" column="0"><widget class="QLabel" name="tl"><property name="text"><string>Up/Down:</string></property></widget></item>
           <item row="0" column="1"><widget class="QLineEdit" name="upper_input"></widget></item>
           <item row="0" column="2"><widget class="QLineEdit" name="lower_input"></widget></item>
           <item row="1" column="0"><widget class="QLabel" name="tr"><property name="text"><string>Left/Right:</string></property></widget></item>
           <item row="1" column="1"><widget class="QLineEdit" name="left_input"></widget></item>
           <item row="1" column="2"><widget class="QLineEdit" name="right_input"></widget></item>
           <item row="0" column="3" rowspan="2"><widget class="QPushButton" name="crop_button"><property name="minimumSize"><size><width>45</width><height>45</height></size></property><property name="text"><string>CROP</string></property></widget></item>
          </layout>
         </item>
         <item>
          <layout class="QHBoxLayout" name="layout_template">
           <item><widget class="QLabel" name="template_label"><property name="text"><string>Template:</string></property></widget></item>
           <item><widget class="QComboBox" name="template_dropdown">
            <property name="minimumSize"><size><width>0</width><height>25</height></size></property>
            <property name="styleSheet"><string>QComboBox { background: #1a1a1e; color: white; border: 1px solid #3a3a40; border-radius: 4px; padding-left: 5px; } QComboBox::drop-down { border: none; }</string></property>
           </widget></item>
          </layout>
         </item>
         <item>
          <layout class="QHBoxLayout" name="layout_bom_sel">
           <item><widget class="QPushButton" name="select_bom_button">
            <property name="minimumSize"><size><width>80</width><height>0</height></size></property>
            <property name="maximumSize"><size><width>80</width><height>16777215</height></size></property>
            <property name="text"><string>BOM</string></property>
           </widget></item>
           <item><widget class="QLineEdit" name="bom_path_display"><property name="placeholderText"><string>Optional: Selected BOM file...</string></property><property name="readOnly"><bool>true</bool></property></widget></item>
          </layout>
         </item>
         <item>
          <layout class="QHBoxLayout" name="layout_pixls_sel">
           <item><widget class="QPushButton" name="select_pixls_button">
            <property name="minimumSize"><size><width>80</width><height>0</height></size></property>
            <property name="maximumSize"><size><width>80</width><height>16777215</height></size></property>
            <property name="text"><string>PIXLs</string></property>
           </widget></item>
           <item><widget class="QLineEdit" name="pix_path_display"><property name="placeholderText"><string>Optional: Selected PIXLs file...</string></property><property name="readOnly"><bool>true</bool></property></widget></item>
          </layout>
         </item>
         <item><widget class="QPushButton" name="generate_document_button"><property name="minimumSize"><size><width>0</width><height>40</height></size></property><property name="font"><font><bold>true</bold></font></property><property name="text"><string>Generate Document</string></property></widget></item>
        </layout>
       </widget>
      </item>
     </layout>
    </item>
    <item>
     <widget class="QFrame" name="footer">
      <property name="minimumSize"><size><width>0</width><height>30</height></size></property>
      <layout class="QHBoxLayout" name="horizontalLayout_footer">
       <item><widget class="QLabel" name="l_copy"><property name="styleSheet"><string>color: #777;</string></property><property name="text"><string>© 2025 Power Integrations</string></property></widget></item>
       <item><spacer><property name="orientation"><enum>Qt::Horizontal</enum></property></spacer></item>
      </layout>
     </widget>
    </item>
   </layout>
  </widget>
 </widget>
 <resources/>
 <connections/>
</ui>