   M S B A M L   `   `   `  ����N  Jcourseworkwarehouse, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullE'clr-namespace:courseworkwarehouse.Formscourseworkwarehouse.Forms  5   	   -    'courseworkwarehouse.Forms.Authorization   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   9local'clr-namespace:courseworkwarehouse.Forms     5   	   _ [MaterialDesignThemes.Wpf, Version=4.5.0.0, Culture=neutral, PublicKeyToken=df2a72020bd7962aFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   -   5          � Language$  ru-RU�5	   	      MaterialDesignBody#���   5
   	   $��Regular"�5   	   $��18(�5   	   &   System.Windows.Media.TextOptions   TextFormattingMode$ Ideal=�5   	      TextRenderingMode$
 Auto=�5   	     MaterialDesignPaper#���  5   	     MaterialDesignFont#���  5   	    � Title$ Авторизация��5   	    � WindowState$ 	Maximized=�5   	    � SizeToContent$ WidthAndHeight=�5   	    � Icon$ /Images/Logo.png��5   	   .��5      � $
��1920��6   $
��1010��5      .+�5   
   '  !MaterialDesignThemes.Wpf.Snackbar  -   5          Name$ Snack��5   "      IsActive . 	   ActionButtonPlacement$	 SeparateLine=�5   "   
 K� ZIndex$
 20��5   "      Message. 5      .  (MaterialDesignThemes.Wpf.SnackbarMessage  $&�� Добро пожаловать!��6-   $��Center=�5   -   $��Center=�5   -   5      �� 5    
   $��Канцелярик��6   $��Center=�5!      $	��Top=�5"      $��72(�5#      ���	#FF005800$��10q�5%      $��Medium"�5&      �� 5'   
   $
��1200��6   $	��800��5(      $��Center=�5)      ���	#FF39B747.+�5+      �� $	��500��6   $	��450��5,      $��/Images/EllipseLogo.png��5-      $��Uniform=�5.      $��10q�5/      �� 51      -   5         ��Name$ loginTextBox��51      $	��650��52      $��20(�53      $��10q�54      )  #MaterialDesignThemes.Wpf.HintAssist   Hint$ 
Логин��55      /  )MaterialDesignThemes.Wpf.TransitionAssist   DisableTransitions . .  (MaterialDesignThemes.Wpf.TextFieldAssist   UnderlineBrush
 �����
������� �� SelectionBrush �	#FF005800 �� 
CaretBrush
 � d � � Cursor$ Arrow��5;      G� 5<      -   5         G�Name$ passwordBox��5<      $	��650��5=      $��20(�5>      $��10q�5?      $ Пароль��5@      
 �����
������� G� SelectionBrush �	#FF005800 G� 
CaretBrush
 � d �$ Arrow��5E      �� 5F      -   5         * &MaterialDesignFlatSecondaryMidBgButton#��[ 5F      $��
Войти��5G      $	��650��5H      $��60��5I      $��26(�5J      ���	#FF005800
����������	#FF39B747$
 Hand��5O      5P      5Q      5R      