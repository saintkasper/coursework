   M S B A M L   `   `   `  ����N  Jcourseworkwarehouse, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:courseworkwarehouse.Formscourseworkwarehouse.Forms  5   	   0    *courseworkwarehouse.Forms.FormEditSupplier   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:courseworkwarehouse.Forms     5   	   _ [MaterialDesignThemes.Wpf, Version=4.5.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	     � Language$  ru-RU�5	   	      MaterialDesignBody#���   5
   	   $��Regular"�5   	   $��18(�5   	   &   System.Windows.Media.TextOptions   TextFormattingMode$ Ideal=�5   	      TextRenderingMode$
 Auto=�5   	     MaterialDesignPaper#���  5   	     MaterialDesignFont#���  5   	    � SizeToContent$ WidthAndHeight=�5   	    � WindowStartupLocation$ CenterScreen=�5   	    � Title$O IДобавление\изменение данных поставщика��5   	    � Icon$ /Images/Logo.png��5   	    � 
ResizeMode$ NoResize=�5   	   .��5      � $
��1000��6   $	��850��5      .+�5   
   �� $
��1000��6   $	��650��5       �� Orientation$ Vertical=�5      $��0,100,0,100q�5      .+�5      ]� $��0,70,0,0q�6   .+�5      �� $��Название��6   $��150,10,0,10q�5      $��18(�5       �� 5!      -   5        	 ��Name$	 tbTitleSupplier��5!      .  (MaterialDesignThemes.Wpf.TextFieldAssist
   UnderlineBrush
 �	#FF39B747 �� SelectionBrush �	#FF005800 �� 
CaretBrush
 � d � �� TextWrapping$
 Wrap=�5%      $��Right=�5&      $��10,10,100,10q�5'      $	��500��5(      $��30��5)      $��18(�5*      5+      ]� 5,      .+�5-      �� $��
Адрес��6   $��150,10,0,10q�5.      $��18(�5/      �� 50      -   5        $	 	tbAddress��50      
 �	#FF39B747 �	#FF005800
 � d �$
 Wrap=�54      $��Right=�55      $��10,10,100,10q�56      $	��500��57      $��30��58      $��18(�59      5:      ]� 5;      .+�5<      �� $!��Номер телефона��6   $��150,10,0,10q�5=      $��18(�5>      �� 5?      -   5        $	 tbPhoneSupplier��5?      
 �	#FF39B747 �	#FF005800
 � d �$
 Wrap=�5C      $��Right=�5D      $��10,10,100,10q�5E      $	��500��5F      $��30��5G       �� 	MaxLength$ 11��5H      $��18(�5I       � Cursor$ Arrow��5J      5K      ]� 5L      .+�5M      �� $��Описание��6   $��150,90,0,10q�5N      $��18(�5O      �� 5P      -   5        $	 tbDescription��5P      
 �	#FF39B747 �	#FF005800
 � d �$
 Wrap=�5T      $��Right=�5U      $	��Top=�5V      $��10,10,100,0q�5W      $	��500��5X      $	��180��5Y      $	 250��5Z      $��18(�5[      5\      �� 5]      -   5         ��Name$ btnSave��5]      $��Сохранить��5^      $	��850��5_      $��0, 100, 0 ,20q�5`      ���	#FF39B747���	#FF39B747$��18(�5d      �� 5e      -   5        $ btnBackSupplier��5e      $��
Назад��5f      $	��850��5g      ���	#FF676767$��18(�5j      5k      5l      5m      