   M S B A M L   `   `   `  ����I  EConversorCofal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null/clr-namespace:ConversorCofalConversorCofal  5   	       ConversorCofal.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   .localclr-namespace:ConversorCofal     5   	     � Title$%  Sicoob Central Cecremge - Cofal��5   	   $	��450��61   $	��800��6>    � 
ResizeMode$ NoResize=�6J    � 	Resources 5	       �� 
TargetType   	TabButton �� Property �� Value �� 
TargetType �� CornerRadius ��Name��%�  &	        �� 5
   
   " ��6   .�5      ��  �@��TextAlignment$ Center=�68   �� 5       � �� 5      �� 5      " ��6&   .R�5      �� $ 2,2,2,2��6!   #��z��5      
�������$��0,1,1,0q�5      ._�5      �� $ contentPresenter��6/   #��z��5   '   #��z��5   '   #��z��5   '   #��z��5   '   #��z��5   '   5      5      5      5      5      .��5       � $	��792��6   $	��419��6   	 � ColumnDefinitions	 5!   
   �� 5"      $
��150* �6   �� 5#      $
��315* �6   .+�5%   
   o� -   5        
 o�Name$

 menu��5%      $
��Left=�6   $��23��68   $	��Top=�6D   $	��792��6\   $��2��6h   .@�5&      m� $��Relatórios��6   $��82��6+    m� Icon 5'      �� 5(      $"��/Assets/Icons/relatorios.png��6   .@�5*      m� -   5         m�Name$" MenuItemRelatorioAposentados��5*   0   $��Aposentados��6    5+      �� 5,      $#��/Assets/Icons/aposentados.png��6    $��24��6G   $��20��6R   5.      m� 5/      -   5        $ MenuItemRelatorioSair��5/   )   $
��sair��6    50      �� 51      $��/Assets/Icons/sair.png��6    $��24��6@   $��16��6K   53      55      56      �� 5:      $��0��6   $��0��6)   $
��auto��66   $
��auto��6D   $��0,70,0,0q�6Q   .+�5;      �� -   5        $��
instrucoes��5;      $��Instruções:��6*   $
��Bold"�6@   $
��Left=�6R   $��0,0,0,0q�6t   5<      �� 5>      $��1��6   $��0��6)   $
��auto��66   $
��auto��6D   $��0,70,0,0q�6Q   .+�5?      �� -   5        $��	passo1Txt��5?      $
��Wrap=�6-   $	��450��6A   $
��Left=�6P   �� 5@      .�6   16   B@ - Selecione o Excel contendo todos os associados que necessite:5A      5B      �� 5C      $��1��6   $��1��6)   $
��auto��66   $
��auto��6D   $��	0,100,0,0q�6Q   .+�5D      �� -   5        $��passo1Button��5D      ���#00AE9D$��30��6C   $	��150��6O   $��75,0,0,0q�6[   $
��Left=�6m   .��5E      �� $
��asap)�6    $
��Bold"�62   Selecionar Arquivo Excel5G      5H      �� 5I      -   5        $��textLblErrorEXCEL��5I      $��0q�6/   $��-140,2,100,0q�6D   $��Center=�6Z   $
��bold"�6w   
���  ��6�   5J      �� 5L      -   5        $��	passo2Txt��5L      $��1��6*   $��2��6:   $
��auto��6G   $
��auto��6U   $��	0,170,0,0q�6b   .+�5M      �� $
��Wrap=�6   $	��450��60   $
��Left=�6?   �� 5N      .�6   26   <: - Selecione o PDF com a relação de todos os associados:5O      5P      �� 5Q      -	   5        $��passo2Button��5Q      $��1��6-   $��1��6=   $
��auto��6J   $
��auto��6X   $��	0,200,0,0q�6e   .+�5R      �� -
   5        $��30��5R      ���#00AE9D$	��150��6:   $��75,0,0,0q�6F   $
��Left=�6t   .��5S      �� $
��asap)�6    $
��Bold"�62   Selecionar Arquivo PDF5U      5V      �� 5W      -   5        $��textLblErrorPDF��5W      $��0q�6-   $��-140,2,100,0q�6B   $��Center=�6X   $
��bold"�6u   
���  ��6�   5X      �� 5Z      -   5        $��	passo3Txt��5Z      $��1��6*   $��2��6:   $
��auto��6G   $
��auto��6U   $��	0,270,0,0q�6b   .+�5[      �� $
��Wrap=�6   $	��450��60   $
��Left=�6?   �� 5\      .�6   36    - Gere o novo arquivo PDF:5]      5^      �� 5_      -   5        $��passo3Button��5_      $��1��6-   $��1��6=   $
��auto��6J   $
��auto��6X   $��	0,300,0,0q�6e   .+�5`      �� -   5        ���#C9D200$��30��5`   /   $	��150��6;   $��75,0,0,0q�6^   $
��Left=�6p   .��5a      �� $
��asap)�6    $
��Bold"�62   	Gerar PDF5c      5d      �� 5e      -   5        $��textLblErrorGerarPDF��5e      $��0q�62   $��-140,2,90,0q�6G   $��Center=�6\   $
��bold"�6y   
���  ��6�   5f      5i      5j      