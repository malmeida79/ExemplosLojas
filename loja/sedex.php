<?php
/**
* @version 1.0 $
* @package SEDEX_Rastrear
* @copyright (C) 2007 Fernando Soares
* <<<<<  Fernando Soares - http://www.fernandosoares.com.br  >>>>>
* @license http://www.gnu.org/copyleft/gpl.html GNU/GPL
*/
 
/** ensure this file is being included by a parent file */
?><title>Rastrear o Sedex</title>
<form action="http://websro.correios.com.br/sro_bin/txect01$.QueryList" method="GET" target="_self">
<INPUT TYPE="hidden" NAME="P_LINGUA" VALUE="001">
<INPUT TYPE="hidden" NAME="P_TIPO" VALUE="001">
<div id="FooterServTitRastr" align="center">											
<div align="left">
Digite o número do SEDEX conforme o exemplo:
<br><input type="text" class="inputbox" name="P_COD_UNI" size="22" maxlength="23" value="SS987654321BR" onFocus="if(this.value=='SS987654321BR')this.value='';" >
<img src="http://www.correios.com.br/images/correios.gif" width=73 height=15 alt="Logo CORREIOS">
<br>
<div align="left">
  <input type="submit" class="button" value="Pesquisar">
</div>
</div>
</div> 
</form>