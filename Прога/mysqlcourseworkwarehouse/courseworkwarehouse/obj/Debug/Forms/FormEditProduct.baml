   M S B A M L   `   `   `  ����N  Jcourseworkwarehouse, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:courseworkwarehouse.Formscourseworkwarehouse.Forms  5   	   /    )courseworkwarehouse.Forms.FormEditProduct   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:courseworkwarehouse.Forms     5   	   _ [MaterialDesignThemes.Wpf, Version=4.5.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	     � Language$  ru-RU�5	   	      MaterialDesignBody#���   5
   	   $��Regular"�5   	   $��18(�5   	   &   System.Windows.Media.TextOptions   TextFormattingMode$ Ideal=�5   	      TextRenderingMode$
 Auto=�5   	     MaterialDesignPaper#���  5   	     MaterialDesignFont#���  5   	    � Title$G AДобавление\изменение данных товара��5   	    � SizeToContent$ WidthAndHeight=�5   	    � WindowStartupLocation$ CenterScreen=�5   	    � 
ResizeMode$ NoResize=�5   	    � Icon$ /Images/Logo.png��5   	   .��5      � $
��1000��6   $	��850��5      .+�5   
   �� $
��1000��6   $	��750��5       �� Orientation$ Vertical=�5      .+�5      ]� $��0,50,0,0q�6   .+�5      �� $��Название��6   $��150,10,0,10q�5      $��18(�5      �� 5      -   5        	 ��Name$	 tbTitleProduct��5      .  (MaterialDesignThemes.Wpf.TextFieldAssist
   UnderlineBrush
 �	#FF39B747 �� SelectionBrush �	#FF005800 �� 
CaretBrush
 � d � �� TextWrapping$
 Wrap=�5#      $��Right=�5$      $��10,10,100,10q�5%      $	��500��5&      $��30��5'      $��18(�5(       � Cursor$ Arrow��5)      5*      ]� 5+      .+�5,      �� $��Цена��6   $��150,10,0,10q�5-      $��18(�5.      �� 5/      -   5        $	 tbCostProduct��5/      
 �	#FF39B747 �	#FF005800
 � d �$
 Wrap=�53      $��Right=�54      $��10,10,100,10q�55      $	��500��56      $��30��57      $��18(�58      59      5:      ]� 5;      .+�5<      �� $��Количество��6   $��150,10,0,10q�5=      $��18(�5>      �� 5?      -   5        $	 tbQuantityProduct��5?      
 �	#FF39B747 �	#FF005800
 � d �$
 Wrap=�5C      $��Right=�5D      $��10,10,100,10q�5E      $	��500��5F      $��30��5G      $��18(�5H      5I      ]� 5J      .+�5K      �� $��Поставщик��6   $��150,10,0,10q�5L      $��18(�5M      �� 5N      -   5         ��Name$ cbSuppliers��5N      $��Bottom=�5O      
 �	#FF39B747$��Right=�5Q      $��10,10,100,10q�5R      $	��500��5S      $��30��5T      $��18(�5U      5V      ]� 5W      .+�5X      �� $#��Товарная группа��6   $��150,10,0,10q�5Y      $��18(�5Z      �� 5[      -   5        $ cbProductGroups��5[      $��Bottom=�5\      
 �	#FF39B747$��Right=�5^      $��10,10,100,10q�5_      $	��500��5`      $��30��5a      $��18(�5b      5c      ]� 5d      .+�5e      �� $��Фото��6   $��
150,80,0,0q�5f      $��18(�5g      �� 5h      ���	#FF005800$��1q�5i      $
��Left=�5j      $	��170��5k      $��
20,10,0,10q�5l      $	��170��5m      ._�5n      �� -   5         ��Name$ 
imagePhoto��5n      $	��170��5o      $	��170��5p      $��Uniform=�5q      5r      �� 5t      -   5        $	 tbPhoto��5t      
 �	#FF39B747 �	#FF005800
 � d �{�.  $
 Wrap=�5y      $��Right=�5z      $	��Top=�5{      $��	10,10,0,0q�5|      $	��500��5}      $	��180��5~      $��18(�5      $ Arrow��5�      �� 5�      -   5         / +MaterialDesignFloatingActionSecondaryButton#��[ 5�      
���  ��
���  ��{�5�      �� *DataContext.ControlsEnabled+ �� RelativeSource �� *FindAncestor+ �� AncestorType$ WindowL�.��5�      '  !MaterialDesignThemes.Wpf.PackIcon     Kind$ DeleteForever=�6.   $��30��5�   .   $��30��5�   .   
�������5�   .   5�      5�      �� 5�      -	   5         ��Name$ btnSelectPhoto��5�      $��Выбрать фото��5�      $	��850��5�      $��0, 80, 0 ,20q�5�      ���	#FF39B747���	#FF39B747$��18(�5�      �� 5�      -
   5        $ btnSaveProduct��5�      $��Сохранить��5�      $	��850��5�      $��0, 0, 0 ,20q�5�      ���	#FF39B747���	#FF39B747$��18(�5�      �� 5�      -   5        $ btnBackProduct��5�      $��
Назад��5�      $	��850��5�      ���	#FF676767$��18(�5�      5�      �� 5�   
   $��Bottom=�6   5�      5�      5�      