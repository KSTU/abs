unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ComCtrls, TabNotBk,math, WordXP, OleServer, ExcelXP;

type
  TForm1 = class(TForm)
    Edit1: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Edit2: TEdit;
    Edit3: TEdit;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Edit4: TEdit;
    Edit5: TEdit;
    Edit6: TEdit;
    Edit7: TEdit;
    Edit8: TEdit;
    Edit9: TEdit;
    Edit12: TEdit;
    Edit13: TEdit;
    Edit14: TEdit;
    Edit15: TEdit;
    Edit16: TEdit;
    Edit17: TEdit;
    Edit18: TEdit;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Edit19: TEdit;
    Edit20: TEdit;
    Edit21: TEdit;
    Label30: TLabel;
    Edit22: TEdit;
    Label31: TLabel;
    Label32: TLabel;
    Edit23: TEdit;
    Edit24: TEdit;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Edit25: TEdit;
    Label39: TLabel;
    Label40: TLabel;
    Label41: TLabel;
    Label42: TLabel;
    Edit26: TEdit;
    Edit27: TEdit;
    Edit28: TEdit;
    Edit29: TEdit;
    Edit31: TEdit;
    Edit32: TEdit;
    Edit33: TEdit;
    Label43: TLabel;
    Button1: TButton;
    ComboBox2: TComboBox;
    Label12: TLabel;
    ComboBox1: TComboBox;
    Label13: TLabel;
    ComboBox3: TComboBox;
    Label44: TLabel;
    ComboBox4: TComboBox;
    ComboBox5: TComboBox;
    ComboBox6: TComboBox;
    ComboBox7: TComboBox;
    ComboBox8: TComboBox;
    ComboBox9: TComboBox;
    Label45: TLabel;
    Label46: TLabel;
    Label47: TLabel;
    Label48: TLabel;
    Label49: TLabel;
    Label50: TLabel;
    Button3: TButton;
    Button4: TButton;
    Button2: TButton;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    Button9: TButton;
    Button10: TButton;
    WordDocument1: TWordDocument;
    Word: TWordApplication;
    Label52: TLabel;
    Edit10: TEdit;
    Label53: TLabel;
    Edit11: TEdit;
    Excel: TExcelApplication;
    Label51: TLabel;
    Label54: TLabel;
    Edit30: TEdit;
    procedure Button1Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox3Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure ComboBox4Change(Sender: TObject);
    procedure ComboBox5Change(Sender: TObject);
    procedure ComboBox6Change(Sender: TObject);
    procedure ComboBox7Change(Sender: TObject);
    procedure ComboBox8Change(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure ComboBox9Change(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Edit1Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  g: real;
  yn:real;
yk:real;
xn:real;
xk:real;
temp:real;
davl:real;
maxdiam:real;
maxh:real;
rog:real;
roa1:real;
roa2:real;
mug:real;
mua1:real;
mua2:real;
mmg:real;
mma1:real;
mma2:real;
mog:real;
moa1:real;
moa2:real;
ras:real;
sig:real;
assic:real;
ga:real;
kg:real;
kn:real;
ce:real;
chas:real;
b0:real;
b1:real;
b2:real;
ce1:real;
nasnaz: array[1..9] of string;
a: array[1..100] of real;
eps: array[1..100] of real;
de: array[1..100] of real;
he: array[1..100] of real;
tip: array[1..100] of integer;
nasa: array[1..100] of real;
nasb: array[1..100] of real;
nasqef:array[1..100] of real;
nasq: array[1..100] of real;
nasp: array[1..100] of real;
nasbn: array[1..100] of real;
comnas: array[1..9] of integer;
standd: array[1..100] of real;
nasup: array[1..100] of integer;
diamin:real;
diaminn:integer;
yob:real;
g0:real;
xravn:real;
lmin,lm,h:real;
linb,dym,dyb,dysr,otl:real;
nasn:integer;
wop,wpr:real;
a2,eps2,de2,he2,nasa2,nasb2,nasq2,nasp2,nasbn2:real;
nasup2,nasqef2:real;
mmsm,rosm,wpr1,wpr2,roy,wpr_o,wpr3:real;
opr1,opr2,opr3:real;
chpi,ob_nul,dia,wrab:real;
S:real;
U:real;
ruGmin:real;
Umin:real;
nprn:integer;
fia:real;
Dy:real;
Rey:real;
Pry:real;
bety:real;
masa1:real;
dpr:real;
Rex:real;
Dx:real;
Prx:real;
betx:real;
Ky:real;
F:real;
Hh:real;
dpsuh:real;
lam:real;
lamt:real;
dpm:real;
kob:integer;
Hh1:real;
Ha1:real;
Hn1:real;
Va:real;
cc1:real;
cc2:real;
cc3:real;
cob:real;
cobmin:real;
vrem:real;
version:string;
//
          minc1: array[1..9] of real;
          minc2: array[1..9] of real;
          minc3: array[1..9] of real;
          minkob: array[1..9] of integer;
          mindia: array[1..9] of real;
          minhh: array[1..9] of real;
          minha: array[1..9] of real;
          minl: array[1..9] of real;
          minw: array[1..9] of real;
          mindp: array[1..9] of real;
          minu: array[1..9] of real;
          minfi: array[1..9] of real;
          minby: array[1..9] of real;
          minbx: array[1..9] of real;
          minky: array[1..9] of real;
          minkparr: array[1..9] of integer;
          minww: array[1..9] of real;
          minll: array[1..9] of real;
          minpsi: array[1..9] of real;
          //
          studl:real;
          studw:real;
          stn:integer;
          nasup1:integer;
          nasqef1:real;
          kolnas:integer;
          prnom: array[1..9] of real;
          masa3:real;
          yn_mol,yk_mol,xn_mol,xk_mol:real;
          yn_mas,yk_mas,xn_mas,xk_mas:real;
          mugnach:real;
          g_ob:real;
          notend:integer;
          kol_parr:integer;
          psi:real;
          zn:integer;
//

implementation

uses Unit2;

{$R *.dfm}
procedure calcdysr();
begin
lm:=lmin*otl/g0;     //"относиетлный" расход абсорбента
xk:=xn+(yn-yk)/lm;   //конечная концентрация в абсорбенте
//linb:=yk-lm*xn;      //это не помню
//xk:=(yn-linb)/lm;
dym:=yk-ras*xn;      //это вроде ок
dyb:=yk+lm*(xk-xn)-ras*xk;  //это тоже вроде ок
if (dym=dyb) then
begin
 dysr:=dym;
end
else
begin
 dysr:=(dyb-dym)/ln(dyb/dym);  //среднелогарифмическая дв. сила
end;
end;

procedure calcwpr();
var i:integer;
begin
wpr_o:=9991;
wpr1:=0.00001;
wpr2:=1.0;
roy:=rosm*273/(273+temp)*davl/0.1013;
while wpr_o>0.001 do
 begin
  wpr3:=wpr1+(wpr2-wpr1)/2.0;
  opr1:=Log10(wpr1*wpr1/9.81/eps2/eps2/eps2*a2*roy/roa2*power(mua2,0.16))-nasa2+nasb2*power(lm,0.25)*power(roy/roa2,0.125);
  opr2:=Log10(wpr2*wpr2/9.81/eps2/eps2/eps2*a2*roy/roa2*power(mua2,0.16))-nasa2+nasb2*power(lm,0.25)*power(roy/roa2,0.125);
  opr3:=Log10(wpr3*wpr3/9.81/eps2/eps2/eps2*a2*roy/roa2*power(mua2,0.16))-nasa2+nasb2*power(lm,0.25)*power(roy/roa2,0.125);
  wpr_o:=abs(opr1-opr2);
  if (opr2>0) and (opr1<0) then
  begin
   if opr3>0 then
   begin
    wpr2:=wpr3;
   end
   else
   begin
   wpr1:=wpr3;
   end;
  end
  else
  begin
   wpr2:=wpr2*2.0;
  end;
 end;
 wpr:=(wpr1+wpr2)/2.0
end;

procedure standartd();
begin
if dia<0.45 then
begin
 dia:=0.4;
end;
if (dia>=0.45) and (dia<0.55) then
begin
 dia:=0.5;
end;
if (dia>=0.55) and (dia<0.65) then
begin
 dia:=0.6;
end;
if (dia>=0.7) and (dia<0.85) then
begin
 dia:=0.8;
end;
if (dia>=0.85) and (dia<1.1) then
begin
 dia:=1.0;
end;
if (dia>=1.1) and (dia<1.3) then
begin
 dia:=1.2;
end;
if (dia>=1.3) and (dia<1.5) then
begin
 dia:=1.4;
end;
if (dia>=1.5) and (dia<1.7) then
begin
 dia:=1.6;
end;
if (dia>=1.7) and (dia<1.9) then
begin
 dia:=1.8;
end;
if (dia>=1.9) and (dia<2.1) then
begin
 dia:=2.0;
end;
if (dia>=2.1) and (dia<2.3) then
begin
 dia:=2.2;
end;
if (dia>=2.3) and (dia<2.5) then
begin
 dia:=2.4;
end;
if (dia>=2.5) and (dia<2.7) then
begin
 dia:=2.6;
end;
if (dia>=2.7) and (dia<2.9) then
begin
 dia:=2.8;
end;
if (dia>=2.9) and (dia<3.1) then
begin
 dia:=3.0;
end;
if (dia>=3.1) and (dia<3.3) then
begin
 dia:=3.2;
end;
if (dia>=3.3) and (dia<3.5) then
begin
 dia:=3.4;
end;
if (dia>=3.5) and (dia<3.7) then
begin
 dia:=3.6;
end;
if (dia>=3.7) and (dia<3.9) then
begin
 dia:=3.8;
end;
if (dia>=3.9) and (dia<4.25) then
begin
 dia:=4.0;
end;
if (dia>=4.25) and (dia<4.75) then
begin
 dia:=4.5;
end;
if (dia>=4.75) and (dia<5.25) then
begin
 dia:=5.0;
end;
if (dia>=5.25) and (dia<5.75) then
begin
 dia:=5.5;
end;
if (dia>=5.75) and (dia<6.25) then
begin
 dia:=6.0;
end;
if (dia>=6.2) and (dia<6.7) then
begin
 dia:=6.4;
end;
if (dia>=6.7) and (dia<7.5) then
begin
 dia:=7.0;
end;
if (dia>=7.5) and (dia<8.5) then
begin
 dia:=8.0;
end;
if (dia>8.5) then
begin
 dia:=9.0;
end;
end;

procedure ardim();
var i:integer;
begin
if diamin<8.0 then diaminn:=27;
if diamin<7.0 then diaminn:=26;
if diamin<6.4 then diaminn:=25;
if diamin<6.0 then diaminn:=24;
if diamin<5.5 then diaminn:=23;
if diamin<5.0 then diaminn:=22;
if diamin<4.5 then diaminn:=21;
if diamin<4.0 then diaminn:=20;
if diamin<3.8 then diaminn:=19;
if diamin<3.6 then diaminn:=18;
if diamin<3.4 then diaminn:=17;
if diamin<3.2 then diaminn:=16;
if diamin<3.0 then diaminn:=15;
if diamin<2.8 then diaminn:=14;
if diamin<2.6 then diaminn:=13;
if diamin<2.4 then diaminn:=12;
if diamin<2.2 then diaminn:=11;
if diamin<2.0 then diaminn:=10;
if diamin<1.8 then diaminn:=9;
if diamin<1.6 then diaminn:=8;
if diamin<1.4 then diaminn:=7;
if diamin<1.2 then diaminn:=6;
if diamin<1.0 then diaminn:=5;
if diamin<0.8 then diaminn:=4;
if diamin<0.6 then diaminn:=3;
if diamin<0.5 then diaminn:=2;
if diamin<0.4 then diaminn:=1;
end;

procedure nasoptim(a1,eps1,de1,he1,nasa1,nasb1,nasq1,nasp1,nasbn1,nasqef1:real;nasup1,nni:integer);
var i,j,k,kol_i:integer;
begin
 a2:=a1;
 eps2:=eps1;
 de2:=de1;
 he2:=he1;
 nasa2:=nasa1;
 nasb2:=nasb1;
 nasq2:=nasq1;
 nasp2:=nasp1;
 nasbn2:=nasbn1;
 nasup2:=nasup1;
 nasqef2:=nasqef1;
 cobmin:=999999999999;
 /////*******************************************************
for kol_i:=1 to 10 do
begin
  kol_parr:=kol_i;
  g:=g_ob;
  g:=g/kol_parr;
  g0:=g*mmsm/22.4*(1.0)/(1.0+yn); //*(1.0-yob)*(mmg/22.4);  //массовый расход газа носителя
  masa1:=g*mmsm/22.4*yn/(1.0+yn);    //массовый расход абсорбтива
  masa3:=masa1-g0*yk;     //масса поглощенного компонента
  lmin:=g0*(yn-yk)/(xravn-xn);   //минимальный расход абсорбента
  for i:=1 to 2500 do
    begin
    otl:=1.0+25/2500*i;
    calcdysr();
    calcwpr();
    ob_nul:=g*(273+temp)/273/davl*0.1013;
    diamin:=sqrt(4.0*ob_nul/chpi/wpr); //минимальны диаметр
    ardim();
      for j:=diaminn to 27 do
         begin
         dia:=standd[j];
         if dia<=maxdiam then
          begin
          wrab:=4.0*ob_nul/chpi/dia/dia; //рабочая скорость
          if (wrab/wpr<0.5) then
            begin
            S:=chpi*dia*dia/4.0; //площадь
            U:=lmin*otl/roa2/S; //L=lmin*otl
            Umin:=a2*nasqef2;
            if U>Umin then   //если плотность орошения больше минимальной
              begin          // считаем дальше
              psi:=1;
              end
            else
              begin
              psi:=0.122*power(U*roa2,1.0/3.0)/sqrt(he2)*power(sig,-0.133/sqrt(he2));
              end;
            if psi>1 then psi:=1.0;
            fia:=3600*U/(a2*(nasp2+3600*nasq2*U));
            Dy:=4.3/100000000*power((temp+273),3.0/2.0)*sqrt(1/mmg+1/mma1)/davl/(((power(mog,1.0/3.0)+power(moa1,1.0/3.0))*(power(mog,1.0/3.0)+power(moa1,1.0/3.0))));
            Rey:=wrab*de2*roy/eps2/(mug/1000);
            Pry:=(mug/1000)/roy/Dy;
            if ((nasup2=1) or (nasup2=2)) then
              begin
              bety:=0.167*Dy/de2*power(rey,0.74)*Power(Pry,0.33)*Power(he2/de2,-0.47);
              end
            else
              begin
              bety:=0.407*power(Rey,0.655)*power(Pry,0.33)*Dy/de2;
              end;
            dpr:=power((mua2/1000)*(mua2/1000)/(roa2*roa2*9.8),1.0/3.0);
            Rex:=4*U*roa2/a2/(mua2/1000);
            Dx:=7.4/power(10,12)*power(assic*(mma2),0.5)*(273+temp)/(mua2)/power(moa1,0.6);
            Prx:=(mua2/1000)/roa2/Dx;
            betx:=0.0021*Dx/dpr*Power(Rex,0.75)*power(Prx,0.5);
            bety:=bety*roy*(1.0/(1.0+yn));
            betx:=betx*roa2;
            Ky:=1/(1/bety+ras/betx);
            F:=masa3/Ky/dysr;  //площадь поверхности
            Hh:=F*4/chpi/a2/dia/dia/fia/psi;
         //гидравлическое сопротивление
            if nasup2=1 then lam:=6.64/power(Rey,0.375);
            if nasup2=4 then lam:=133.0/Rey+2.34;
            if nasup2=3 then
            begin
              if (Rey<40) then lam:=140.0/Rey else lam:=16.0/power(Rey,0.2);
            end;
            if nasup2=2 then lam:=lamt+(4.2/eps2/eps2-8.1/eps2+3.9)*de2/he2;
            dpsuh:=lam*Hh/de2*wrab*wrab/eps2/eps2/2*roy;
            dpm:=dpsuh*power(10,nasbn2*U);
         //экономическая часть
            kob:=trunc(Hh/maxh)+1;
            Hh1:=Hh/kob; //высота одного абсорбера
            Hn1:=hh1+0.3*(hh1/25/0.3-1); //высота насадочной части
            Ha1:=hn1+1.05*dia+2.4;   //высота всего аппарата
            Va:=Ha1*Chpi*dia*dia/4.0;   //объем аппарата

            cc1:=ga*(b0+b1*Va+b2*Hn1*dia*Chpi)*kob*kol_parr/100;
            cc2:=0.1*dpm*g0/22.4*mmsm*chas*ce*kob*kol_parr/(kg*rosm);
            cc3:=(0.1*9.8*Ha1*ce/kn+3600*ce1)*lm*g0*kol_parr*chas;
            cob:=cc1+cc2+cc3;      //суммарные затраты
            if cob<cobmin then
              begin
              cobmin:=cob;        //оптимальные значения
              minkparr[nni]:=kol_parr;
              minww[nni]:=wrab/wpr;
              minll[nni]:=otl;
              minc1[nni]:=cc1;
              minc2[nni]:=cc2;
              minc3[nni]:=cc3;
              minkob[nni]:=kob;
              mindia[nni]:=dia;
              minhh[nni]:=hh1;
              minha[nni]:=ha1;
              minl[nni]:=lm*g0;
              minw[nni]:=wrab;
              mindp[nni]:=dpm;
              minu[nni]:=U;
              minfi[nni]:=fia;
              minby[nni]:=bety;
              minbx[nni]:=betx;
              minky[nni]:=Ky;
              minpsi[nni]:=psi;
              end;
            end;
         end;
      end;
    end
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var i:integer;
begin
//showmessage(inttostr(trunc(1)));
lamt:=0.053;
chpi:=3.1415926535897932384626433832795;
g:=strtofloat(edit1.Text);
yn:=strtofloat(edit2.Text);
yk:=strtofloat(edit3.Text);
xn:=strtofloat(edit4.Text);
temp:=strtofloat(edit6.Text);
davl:=strtofloat(edit7.Text);
maxdiam:=strtofloat(edit8.Text);
maxh:=strtofloat(edit9.Text);
studl:=strtofloat(edit10.Text);
studw:=strtofloat(edit11.Text);
roa2:=strtofloat(edit12.Text);
mug:=strtofloat(edit13.Text);
mugnach:=mug;
mua1:=strtofloat(edit14.Text);
mua2:=strtofloat(edit15.Text);
rog:=mmg/22.4*davl/0.1013;
roa1:=mma1/22.4*davl/0.1013;
mmg:=strtofloat(edit16.Text);
mma1:=strtofloat(edit17.Text);
mma2:=strtofloat(edit18.Text);
mog:=strtofloat(edit19.Text);
moa1:=strtofloat(edit20.Text);
moa2:=strtofloat(edit21.Text);
ras:=strtofloat(edit22.Text);
sig:=strtofloat(edit23.Text);
assic:=strtofloat(edit24.Text);
ga:=strtofloat(edit25.Text);
kg:=strtofloat(edit26.Text);
kn:=strtofloat(edit27.Text);
ce:=strtofloat(edit28.Text);
chas:=strtofloat(edit29.Text);
b0:=strtofloat(edit31.Text);
b1:=strtofloat(edit32.Text);
b2:=strtofloat(edit30.Text);
ce1:=strtofloat(edit33.Text);

//РАСЧЕТ
mmsm:=(yn+1.0)/(yn/mma1+1.0/mmg);  //молярная масса газовой смеси
yn_mol:=(yn/mma1)/(yn/mma1+(1.0)/mmg);
yn_mas:=yn/(1+yn);
yk_mol:=(yk/mma1)/(yk/mma1+(1.0)/mmg);
yk_mas:=yk/(1+yk);
mug:=power(10,yn_mol*log10(mua1)+(1.0-yn_mol)*log10(mug));
xn_mol:=(xn/mma1)/(xn/mma1+(1.0)/mma2);
xn_mas:=xn/(1+xn);
rosm:=mmsm/22.4;    //плотность газовой смеси
//считываем параметры насадки

//CALCULATION START
 yob:=yn*mmg/mma1;    //объемная/мольная доля абсорбтива
 g_ob:=g;           //pfgjvbyftv hfc[jl
 g0:=g*mmsm/22.4*(1.0)/(1.0+yn); //*(1.0-yob)*(mmg/22.4);  //массовый расход газа носителя
 masa1:=g*mmsm/22.4*yn/(1.0+yn);    //массовый расход абсорбтива
 masa3:=masa1-g0*yk;
 xravn:=yn/ras;     //равновесная концентрация в жидкости
 lmin:=g0*(yn-yk)/(xravn-xn);   //минимальный расход абсорбента
 nasn:=0;
 //Далее идет перебор насадок
 //вставляем номера
 nasnaz[1]:=unit1.form1.combobox1.text;
 nasnaz[2]:=unit1.form1.combobox2.text;
 nasnaz[3]:=unit1.form1.combobox3.text;
 nasnaz[4]:=unit1.form1.combobox4.text;
 nasnaz[5]:=unit1.form1.combobox5.text;
 nasnaz[6]:=unit1.form1.combobox6.text;
 nasnaz[7]:=unit1.form1.combobox7.text;
 nasnaz[8]:=unit1.form1.combobox8.text;
 nasnaz[9]:=unit1.form1.combobox9.text;
 if unit1.Form1.ComboBox1.ItemIndex=28 then
 comnas[1]:=51
 else
 comnas[1]:=unit1.Form1.ComboBox1.ItemIndex+1;
//
 if unit1.Form1.ComboBox2.ItemIndex=28 then
 comnas[2]:=52
 else
 comnas[2]:=unit1.Form1.ComboBox2.ItemIndex+1;
//
 if unit1.Form1.ComboBox3.ItemIndex=28 then
 comnas[3]:=53
 else
 comnas[3]:=unit1.Form1.ComboBox3.ItemIndex+1;
//
 if unit1.Form1.ComboBox4.ItemIndex=28 then
 comnas[4]:=54
 else
 comnas[4]:=unit1.Form1.ComboBox4.ItemIndex+1;
//
 if unit1.Form1.ComboBox5.ItemIndex=28 then
 comnas[5]:=55
 else
 comnas[5]:=unit1.Form1.ComboBox5.ItemIndex+1;
//
 if unit1.Form1.ComboBox6.ItemIndex=28 then
 comnas[6]:=56
 else
 comnas[6]:=unit1.Form1.ComboBox6.ItemIndex+1;
//
 if unit1.Form1.ComboBox7.ItemIndex=28 then
 comnas[7]:=57
 else
 comnas[7]:=unit1.Form1.ComboBox7.ItemIndex+1;
//
 if unit1.Form1.ComboBox8.ItemIndex=28 then
 comnas[8]:=58
 else
 comnas[8]:=unit1.Form1.ComboBox8.ItemIndex+1;
//
 if unit1.Form1.ComboBox9.ItemIndex=28 then
 comnas[9]:=59
 else
 comnas[9]:=unit1.Form1.ComboBox9.ItemIndex+1;
//
 for i:=1 to 9 do
 begin
  nprn:=comnas[i];  //номер насадки в массивах
  //showmessage(inttostr(nprn));
    if ((nprn<>35) and (nprn<>0)) then
    begin
      nasn:=nasn+1;     //номер насадки по очереди
    //оптимизация насадки
      nasoptim(a[nprn],eps[nprn],de[nprn],he[nprn],nasa[nprn],nasb[nprn],nasq[nprn],nasp[nprn],nasbn[nprn],nasqef[nprn],nasup[nprn],i);
    end;
 end;
 //технологический расчет при заданных условиях
 stn:=comnas[1];    //номер насадки студента
 a2:=a[stn];
 eps2:=eps[stn];
 de2:=de[stn];
 he2:=he[stn];
 nasa2:=nasa[stn];
 nasb2:=nasb[stn];
 nasq2:=nasq[stn];
 nasp2:=nasp[stn];
 nasbn2:=nasbn[stn];
 nasup2:=nasup[stn];
 nasqef2:=nasqef[stn];
 /////*******************************************************
 notend:=1;
 kol_parr:=1;
 while notend=1 do
begin
 g:=g_ob;
 g:=g/kol_parr;
 g0:=g*mmsm/22.4*(1.0)/(1.0+yn); //*(1.0-yob)*(mmg/22.4);  //массовый расход газа носителя
 masa1:=g*mmsm/22.4*yn/(1.0+yn);    //массовый расход абсорбтива
 masa3:=masa1-g0*yk;     //масса поглощенного компонента
 //xravn:=yn/ras;     //равновесная концентрация в жидкости
 lmin:=g0*(yn-yk)/(xravn-xn);   //минимальный расход абсорбента
 /// вставили сверху
  otl:=studl;
  calcdysr();
  calcwpr();
  //оптимизируем w
  ob_nul:=g*(273+temp)/273/davl*0.1013;
  diamin:=sqrt(4.0*ob_nul/chpi/wpr); //минимальны диаметр
  //standartd(); //стандартизируем диаметр
  ardim();
  //showmessage(floattostr(diamin)+ '|' +floattostr(diaminn) + '|' + floattostr(standd[diaminn]));
  if standd[diaminn]<=maxdiam then
   begin
   notend:=0;
   end
  else
   begin
    kol_parr:=kol_parr+1;
   end;
end;
 //showmessage(inttostr(kol_parr));
 ////********************************************************

 calcdysr();
 calcwpr();
 ob_nul:=g*(273+temp)/273/davl*0.1013;
 wrab:=studw*wpr;
 diamin:=sqrt(4.0*ob_nul/chpi/wrab); //минимальны диаметр
 //standartd(); //стандартизируем диаметр
 ardim();
 dia:=standd[diaminn];
 wrab:=4.0*ob_nul/chpi/dia/dia; //рабочая скорость
 S:=chpi*dia*dia/4.0; //площадь
 U:=lmin*studl/roa2/S; //L=lmin*otl
 Umin:=a2*nasqef2;
 if U>Umin then   //если плотность орошения больше минимальной
          begin          // считаем дальше
          psi:=1;         end
        else
          begin
          psi:=0.122*power(U*roa2,1.0/3.0)/sqrt(he2)*power(sig,-0.133/sqrt(he2));
          end;
          if psi>1 then psi:=1.0;
        fia:=3600*U/(a2*(nasp2+3600*nasq2*U));
        //расчет коэффициентов массотдачи
         //vrem:=((power(mog,1.0/3.0)+power(moa1,1.0/3.0))*(power(mog,1.0/3.0)+power(moa1,1.0/3.0)));
         Dy:=4.3/100000000*power((temp+273),3.0/2.0)*sqrt(1/mmg+1/mma1)/davl/(((power(mog,1.0/3.0)+power(moa1,1.0/3.0))*(power(mog,1.0/3.0)+power(moa1,1.0/3.0))));
         Rey:=wrab*de2*roy/eps2/(mug/1000);
         Pry:=(mug/1000)/roy/Dy;
         if ((nasup2=1) or (nasup2=2)) then
          begin
          bety:=0.167*Dy/de2*power(rey,0.74)*Power(Pry,0.33)*Power(he2/de2,-0.47);
          end
          else
          begin
          bety:=0.407*power(Rey,0.655)*power(Pry,0.33)*Dy/de2;
          end;
         dpr:=power((mua2/1000)*(mua2/1000)/(roa2*roa2*9.8),1.0/3.0);
         Rex:=4*U*roa2/a2/(mua2/1000);
         Dx:=7.4/power(10,12)*power(assic*(mma2),0.5)*(273+temp)/(mua2)/power(moa1,0.6);
         Prx:=(mua2/1000)/roa2/Dx;
         betx:=0.0021*Dx/dpr*Power(Rex,0.75)*power(Prx,0.5);
         bety:=bety*roy*(1.0/(1.0+yn));
         betx:=betx*roa2;
         Ky:=1/(1/bety+ras/betx);
         F:=masa3/Ky/dysr;  //площадь поверхности
         Hh:=F*4/chpi/a2/dia/dia/fia/psi;
         //гидравлическое сопротивление
         if nasup2=1 then lam:=6.64/power(Rey,0.375);
         if nasup2=4 then lam:=133.0/Rey+2.34;
         if nasup2=3 then
            begin
            if (Rey<40) then lam:=140.0/Rey else lam:=16.0/power(Rey,0.2);
            end;
         if nasup2=2 then lam:=lamt+(4.2/eps2/eps2-8.1/eps2+3.9)*de2/he2;
         dpsuh:=lam*Hh/de2*wrab*wrab/eps2/eps2/2*roy;
         dpm:=dpsuh*power(10,nasbn2*U);
         //экономическая часть
         kob:=trunc(Hh/maxh)+1;
         Hh1:=Hh/kob; //высота одного абсорбера
         Hn1:=hh1+0.3*(hh1/25/0.3-1); //высота насадочной части
         Ha1:=hn1+1.05*dia+2.4;   //высота всего аппарата
         Va:=Ha1*Chpi*dia*dia/4.0;   //объем аппарата

            cc1:=ga*(b0+b1*Va+b2*Hn1*dia*Chpi)*kob*kol_parr/100;
            cc2:=0.1*dpm*g0/22.4*mmsm*chas*ce*kob*kol_parr/(kg*rosm);
            cc3:=(0.1*9.8*Ha1*ce/kn+3600*ce1)*lm*g0*kol_parr*chas;
            cob:=cc1+cc2+cc3;      //суммарные затраты
 //end
// else
//      begin
//      //showmessage('U<Umin. Необходимо или увеличить расход абсорбента или уменьшить площадь орошения')
//      end;

 //вывод
 word.Connect;
 word.Visible:=false;
 word.Documents.Add(EmptyParam,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.PageSetup.TopMargin:=30;
 word.Selection.PageSetup.LeftMargin:=25;
 word.Selection.PageSetup.RightMargin:=25;
 word.Selection.PageSetup.BottomMargin:=30;
 word.Selection.Font.Size:=6;
 word.Selection.TypeText('Версия программы:'+version);
 word.Selection.TypeParagraph;
 word.Selection.Font.Size:=10;
 word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeText('Исходные данные:');
 word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeParagraph;
 word.Selection.Font.Size:=8;
 word.Selection.TypeText('Расход газовой смеси (н.у.): '+floattostr(g*kol_parr)+ ' м');
  word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('/с         (Массовый расход: '+floattostr(trunc(g*mmsm/22.4*10000)/10000)+ ' кг/с)');
  word.Selection.TypeParagraph;
 word.Selection.TypeText('Концентрации абсорбтива в газе, кгА/кгY:');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('начальная (');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('y');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('н');
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('): '+floattostr(yn) + '    ');
 word.Selection.TypeText('(мольных %: '+floattostr(trunc(yn_mol*100000)/1000) + ')   ');
 word.Selection.TypeText('(массовых %: '+floattostr(trunc(yn_mas*100000)/1000) + ')   ');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('конечная (');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('y');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('к');
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('): '+floattostr(yk) + '   ');
 word.Selection.TypeText('(мольных %: '+floattostr(trunc(yk_mol*100000)/1000) + ')   ');
  word.Selection.TypeText('(массовых %: '+floattostr(trunc(yk_mas*100000)/1000) + ')   ');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Концентрации абсорбтива в жидкости, кгA/кгY');
 word.Selection.TypeParagraph;

 word.Selection.TypeText('начальная (');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('x');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('н');
 word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('): '+floattostr(xn));

  word.Selection.TypeText('(мольных %: '+floattostr(trunc(xn_mol*100000)/1000) + ')   ');
  word.Selection.TypeText('(массовых %: '+floattostr(trunc(xn_mas*100000)/1000) + ')   ');
 word.Selection.TypeParagraph;
 //word.Selection.TypeText('конечная (xk): '+floattostr(trunc(xk*1000000)/1000000));
 //word.Selection.TypeParagraph;
 word.Selection.TypeText('Температура в аппарате: '+ floattostr(temp)+' ');
 word.Selection.InsertSymbol(176,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText(' C');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Давление в аппарате: '+floattostr(davl)+' МПа');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Максимальный диаметр аппарата: '+floattostr(maxdiam)+' м');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Максимальная высота аппарата: ' + floattostr(maxh)+ ' м');
 word.Selection.TypeParagraph;
 word.Selection.Font.Bold:=wdToggle;
    word.Selection.TypeText('Теплофизические свойства:');
    word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Плотность абсорбента: '+floattostr(roa2)+' кг/м');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Молярная масса инертного газа: '+floattostr(mmg)+' г/моль      ');
 word.Selection.TypeText('Мольный объем инертного газа: '+floattostr(mog)+' см');
  word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('/моль');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Молярная масса абсорбтива: '+floattostr(mma1)+' г/моль         ');
  word.Selection.TypeText('Мольный объем абсорбтива: '+floattostr(moa1)+' см');
  word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('/моль');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Молярная масса абсорбента: '+floattostr(mma2)+' г/моль         ');
  //word.Selection.TypeText('Мольный объем абсорбента: '+floattostr(moa2)+' см');
 // word.Selection.font.Superscript:=wdToggle;
 //word.Selection.TypeText('3');
 //word.Selection.font.Superscript:=wdToggle;
 //word.Selection.TypeText('/моль');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Коэффициент вязкости инертного газа: '+floattostr(mugnach)+' мПа');
 word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText('с');
  word.Selection.TypeParagraph;
 word.Selection.TypeText('Коэффициент вязкости абсорбтива: '+floattostr(mua1)+' мПа');
  word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText('с');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Коэффициент вязкости абсорбента: '+floattostr(mua2)+' мПа');
  word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText('с');
  word.Selection.TypeParagraph;
    word.Selection.TypeText('Коэффициент вязкости газовой смеси: '+floattostr(trunc(mug*100000)/100000)+' мПа');
  word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText('с');
  word.Selection.TypeParagraph;
 word.Selection.TypeText('Коэффициент распределения: '+floattostr(ras)+' кг/кг');
 word.Selection.TypeParagraph;
  word.Selection.TypeText('Коэффициент поверхностного натяжения: '+floattostr(sig)+' Н/м');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Параметр, учитывающий ассоциацию молекул: '+floattostr(assic));
 word.Selection.TypeParagraph;
 word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeText('Экономические показатели:');
 word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Годовые отчисления от стоимости аппарата:' +floattostr(ga)+ ' %');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('КПД газодувки: ' +floattostr(kg)+ ' %');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('КПД насоса: ' +floattostr(kn)+ ' %');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Цена электроэнергии: ' +floattostr(ce)+ ' у.е./кВт');
 word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
 word.Selection.TypeText('ч');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Работа абсорбера в год: ' +floattostr(chas)+ ' часов');
 word.Selection.TypeParagraph;
 word.Selection.TypeText('Коэффициент стоимости монтажа аппарата ');
  word.Selection.font.Italic:=wdToggle;
 word.Selection.TypeText('B');
 word.Selection.font.Italic:=wdToggle;
 word.Selection.TypeText('0: ' +floattostr(b0)+ ' у.е.');

 word.Selection.TypeParagraph;
  word.Selection.TypeText('Коэффициент стоимости материала ');
   word.Selection.font.Italic:=wdToggle;
 word.Selection.TypeText('B');
 word.Selection.font.Italic:=wdToggle;
   word.Selection.TypeText('1: ' +floattostr(b1)+ ' у.е./м');

   word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeParagraph;

   word.Selection.TypeText('Коэффициент стоимости материала ');
   word.Selection.font.Italic:=wdToggle;
 word.Selection.TypeText('B');
 word.Selection.font.Italic:=wdToggle;
   word.Selection.TypeText('2: ' +floattostr(b2)+ ' у.е./м');
      word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeParagraph;
   word.Selection.TypeText('Затраты на регенерацию абсорбента: ' +floattostr(ce1)+ ' у.е./кг');
 word.Selection.TypeParagraph;
 word.Selection.TypeParagraph;
 word.Selection.Font.Bold:=wdToggle;
    word.Selection.TypeText('Технологический расчет варианта студента');
    word.Selection.Font.Bold:=wdToggle;
 word.Selection.TypeParagraph;
    word.Selection.TypeText('Отношение ');
    word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('L');
    word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('/');
    word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('L');
    word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('мин: ' +floattostr(studl));
 word.Selection.TypeParagraph;
     word.Selection.TypeText('Отношение ');
     word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('W');
    word.Selection.Font.Italic:=wdToggle;
     word.Selection.TypeText('/');
     word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('W');
    word.Selection.Font.Italic:=wdToggle;
     word.Selection.TypeText('пр: ' +floattostr(studw));
 word.Selection.TypeParagraph;
      word.Selection.TypeText('Параметры насадки:');
 word.Selection.TypeParagraph;
   word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('a');
 word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText(': ' +floattostr(a[stn])+ ' м');
     word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
     word.Selection.TypeText('/м');
          word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeParagraph;
 word.Selection.TypeText(' ');
   word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('e');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText(': ' +floattostr(eps[stn])+ ' м');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
     word.Selection.TypeText('/м');
          word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
  word.Selection.TypeParagraph;
  word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('A');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText(': ' +floattostr(nasa[stn]));
  word.Selection.TypeParagraph;
  word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('B');
 word.Selection.Font.Italic:=wdToggle;
  word.Selection.TypeText(': ' +floattostr(nasb[stn]));
  word.Selection.TypeParagraph;
  word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('p');
 word.Selection.Font.Italic:=wdToggle;
   word.Selection.TypeText(': ' +floattostr(nasp[stn]));
  word.Selection.TypeParagraph;
  word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('q');
 word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText(': ' +floattostr(nasq[stn]));
  word.Selection.TypeParagraph;
  word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('b');
 word.Selection.Font.Italic:=wdToggle;
  word.Selection.TypeText(': ' +floattostr(nasbn[stn]));
  word.Selection.TypeParagraph;
  word.Selection.Font.Bold:=wdToggle;
  word.Selection.TypeText('Расчет:');
  word.Selection.Font.Bold:=wdToggle;
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Расход газовой смеси на параллельное подключение (н.у.): '+floattostr(g)+ ' м');
  word.Selection.font.Superscript:=wdToggle;
  word.Selection.TypeText('3');
  word.Selection.font.Superscript:=wdToggle;
  word.Selection.TypeText('/с');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Расход абсорбента: ' +floattostr(trunc(lm*g0*10000)/10000)+ ' кг/с   ');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('(Минимальный расход абсорбента: ' +floattostr(trunc(lmin*10000)/10000)+ ' кг/с)   ');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Конечная концентрация в абсорбенте (');
    word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('x');
 word.Selection.Font.Italic:=wdToggle;
   word.Selection.font.Subscript:=wdToggle;
 word.Selection.TypeText('к');
 word.Selection.font.Subscript:=wdToggle;
  word.Selection.TypeText('): ' +floattostr(trunc(xk*1000000)/1000000)+ ' кгA/кгY');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Средняя движущая сила : ' +floattostr(trunc(dysr*1000000)/1000000)+ ' кгA/кгY');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Предельная скорость газа : ' +floattostr(trunc(wpr*1000000)/1000000)+ ' м/с');
  word.Selection.TypeParagraph;
    word.Selection.TypeText('Диаметр абсорбера : ' +floattostr(trunc(dia*1000000)/1000000)+ ' м');
  word.Selection.TypeParagraph;
    word.Selection.TypeText('Рабочая скорость газа : ' +floattostr(trunc(wrab*1000000)/1000000)+ ' м/с');
  word.Selection.TypeParagraph;
 word.Selection.TypeText('Плотность орошения : ' +floattostr(trunc(U*1000000)/1000000)+ ' м');
 word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('3');
 word.Selection.font.Superscript:=wdToggle;
  word.Selection.TypeText('/(м');
  word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
  word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
   word.Selection.TypeText('с)');
  word.Selection.TypeParagraph;
//
if (U<Umin) then
  begin
  word.Selection.TypeText('(');
    word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('U');
 word.Selection.Font.Italic:=wdToggle;
  word.Selection.TypeText('<');
   word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('U');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('min) Коэффициент смачиваемости насадки : ' +floattostr(trunc(psi*1000000)/1000000));

  word.Selection.TypeParagraph;
  end
else
  begin
  word.Selection.TypeText('(');
    word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('U');
 word.Selection.Font.Italic:=wdToggle;
  word.Selection.TypeText('>');
   word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('U');
 word.Selection.Font.Italic:=wdToggle;
 word.Selection.TypeText('min) Коэффициент смачиваемости насадки : ' +floattostr(trunc(psi*1000000)/1000000));
  word.Selection.TypeParagraph;
  end;
//
   word.Selection.TypeText('Доля активной поверхности насадки : ' +floattostr(trunc(fia*1000000)/1000000));
  word.Selection.TypeParagraph;
    word.Selection.TypeText('Коэффициент массоотдачи в газовой фазе : ' +floattostr(trunc(bety*1000000)/1000000)+ ' кг/(м');
    word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
   word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
   word.Selection.TypeText('с)');
  word.Selection.TypeParagraph;
     word.Selection.TypeText('Коэффициент массоотдачи в жидкой фазе : ' +floattostr(trunc(betx*1000000)/1000000)+ ' кг/(м');
         word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
   word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
   word.Selection.TypeText('с)');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Коэффициент масспередачи : ' +floattostr(trunc(Ky*1000000)/1000000)+ ' кг/(м');
  word.Selection.font.Superscript:=wdToggle;
 word.Selection.TypeText('2');
 word.Selection.font.Superscript:=wdToggle;
   word.Selection.InsertSymbol(183,EmptyParam,EmptyParam,EmptyParam);
   word.Selection.TypeText('с)');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Поверхность масспередачи : ' +floattostr(trunc(F*1000000)/1000000)+ ' м');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Требуемая высота : ' +floattostr(trunc(Hh*1000000)/1000000)+ ' м');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Количество абсорберов установленных последовательно: ' +inttostr(kob));
  // word.Selection.TypeText('Количество абсорберов установленных последовательно: ' + tostr(kob));
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Количество абсорберов установленных параллельно: ' +inttostr(kol_parr));
  word.Selection.TypeParagraph;
    word.Selection.TypeText('Высота насадочной части одного абсорбера : ' +floattostr(trunc(Hh1*1000000)/1000000)+ ' м');
  word.Selection.TypeParagraph;
     word.Selection.TypeText('Высота одного абсорбера : ' +floattostr(trunc(Ha1*1000000)/1000000)+ ' м');
  word.Selection.TypeParagraph;
  if (dpm<100000) then
  begin
     word.Selection.TypeText('Гидравлическое сопротивление : ' +floattostr(trunc(dpm*1000000)/1000000)+ ' Па');
  end
  else
  begin
     word.Selection.TypeText('Гидравлическое сопротивление : ' +floattostr(dpm)+ ' Па');
  end;
     //word.Selection.TypeText('Гидравлическое сопротивление : ' +floattostr(dpm)+ ' Па');
  word.Selection.TypeParagraph;
       word.Selection.TypeText('Амортизационные отчисления  : ' + floattostr(trunc(cc1*1000000)/1000000)+ ' у.е./год');
  word.Selection.TypeParagraph;
  if (cc2<100000) then
  begin
    word.Selection.TypeText('Затраты на прокачку газа  : ' +floattostr(trunc(cc2*1000000)/1000000)+ ' у.е./год');
  end
  else
  begin
     word.Selection.TypeText('Затраты на прокачку газа  : ' +floattostr(cc2)+ ' у.е./год');
  end;
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Затраты на перекачку и регенерацию абсорбента  : ' +floattostr(trunc(cc3*1000000)/1000000)+ ' у.е./год');
  word.Selection.TypeParagraph;
  word.Selection.TypeText('Суммарные затраты  : ' +floattostr(trunc(cob*1000000)/1000000)+ ' у.е./год');
  word.Selection.TypeParagraph;
  //вывод оптимизации
//создаем таблицу
word.Selection.InsertBreak(EmptyParam);
word.Selection.Font.Bold:=wdToggle;
word.Selection.TypeText('Оптимизация');
word.Selection.Font.Bold:=wdToggle;
word.Options.DefaultBorderLineStyle:=wdLineStyleSingle;
word.Selection.Tables.Add(word.Selection.Range,22,nasn+1,EmptyParam,EmptyParam);
word.Selection.TypeText('Параметры используемой насадки');word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    if comnas[i]=28 then
    begin
    word.Selection.TypeText('a '+floattostr(a[comnas[i]])+' | e '+floattostr(eps[comnas[i]])+' | de '+floattostr(de[comnas[i]])+' | A'+floattostr(nasa[comnas[i]])+' | B'+floattostr(nasb[comnas[i]])+' | p'+floattostr(nasb[comnas[i]])+' | q'+floattostr(nasq[comnas[i]])+' | b'+floattostr(nasb[comnas[i]]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    end
    else
    begin
    word.Selection.TypeText(nasnaz[i]);
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    end;
    //выводимое свойство
    end;
 end;
 word.Selection.TypeText('Диаметр абсорбера, м');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(mindia[i]*100+0.2)/100));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
  word.Selection.TypeText('Количество абсорберов соединенных параллельно');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(inttostr(minkparr[i]));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
  word.Selection.TypeText('Количество абсорберов соединенных последовательно');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(inttostr(minkob[i]));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
word.Selection.TypeText('Общее количество абсорберов');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(inttostr(minkob[i]*minkparr[i]));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
 word.Selection.TypeText('Высота насадки одного абсорбера, м');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minhh[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
 //
  word.Selection.TypeText('Высота одного абсорбера, м');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minha[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
 //
  word.Selection.TypeText('Расход абсорбента, кг/с');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minl[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
  word.Selection.Font.Italic:=wdToggle;
  word.Selection.TypeText('L/Lmin');
  word.Selection.Font.Italic:=wdToggle;
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minll[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
   word.Selection.TypeText('Рабочая фиктивная скорость, м/с');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minw[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
 word.Selection.Font.Italic:=wdToggle;
    word.Selection.TypeText('W/Wпр');
    word.Selection.Font.Italic:=wdToggle;
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minww[i]*10000)/10000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
    word.Selection.TypeText('Плотность орошения, куб.м/(кв.м с)');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minu[i]*1000000)/1000000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
//
    word.Selection.TypeText('Смачиваемость насадки');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minpsi[i]*1000000)/1000000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
//
     word.Selection.TypeText('Доля активной поверхности');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minfi[i]*100000)/100000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
      word.Selection.TypeText('Коэффициент массоотдачи в газовой фазе, кг/(кв.м с)');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minby[i]*10000000)/10000000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
       word.Selection.TypeText('Коэффициент массоотдачи в жидкой фазе, кг/(кв.м с)');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minbx[i]*10000000)/10000000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
        word.Selection.TypeText('Коэффициент массопередачи, кг/(кв.м с)');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minky[i]*10000000)/10000000));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
 word.Selection.TypeText('Гидравлическое сопротивление, Па');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(mindp[i]*100)/100));
    // word.Selection.TypeText(floattostr(mindp[i]));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
  word.Selection.TypeText('Амортизационные отчисления, у.е./год');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minc1[i]*100)/100));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
   word.Selection.TypeText('Затраты на прокачку газа, у.е./год');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minc2[i]*100)/100));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
    word.Selection.TypeText('Затраты на перекачку и регенерацию абсорбента, у.е./год');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc(minc3[i]*100)/100));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
     word.Selection.TypeText('Суммарные затраты, у.е./год');
word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
for i:=1 to 9 do
 begin
    if ((comnas[i]<>33) and (comnas[i]<>0)) then
    begin
    word.Selection.TypeText(floattostr(trunc((minc1[i]+minc2[i]+minc3[i])*100)/100));
    //showmessage(floattostr(mindia[i]));
    word.Selection.MoveRight(EmptyParam,EmptyParam,EmptyParam);
    //выводимое свойство
    end;
 end;
  //
    word.Visible:=true;
    word.PrintOut;
    word.ActiveDocument.Saved:=true;
    word.ActiveDocument.Close(EmptyParam,EmptyParam,EmptyParam);    word.Quit;
    showmessage('Документ с результатами отослан на принтер установленный по умолчанию');
    close;

end;



procedure TForm1.ComboBox1Change(Sender: TObject);
begin
button4.Visible:=true;
combobox2.Enabled:=true;
//showmessage(inttostr(combobox1.ItemIndex));
//if combobox1.ItemIndex=3 then
//  begin
//  button4.Visible:=false;
//  end;
//showmessage(inttostr(combobox1.ItemIndex));
end;

procedure TForm1.ComboBox3Change(Sender: TObject);
begin
combobox4.Enabled:=true;
button2.Visible:=true;
if combobox3.ItemIndex=32 then
  begin
  button2.Visible:=false;
  end;
end;

procedure TForm1.ComboBox2Change(Sender: TObject);
begin
combobox3.Enabled:=true;
button3.Visible:=true;
if combobox2.ItemIndex=32 then
  begin
  button3.Visible:=false;
  end;
end;

procedure TForm1.ComboBox4Change(Sender: TObject);
begin
combobox5.Enabled:=true;
button5.Visible:=true;
if combobox4.ItemIndex=32 then
  begin
  button5.Visible:=false;
  end;
end;

procedure TForm1.ComboBox5Change(Sender: TObject);
begin
combobox6.Enabled:=true;
button6.Visible:=true;
if combobox5.ItemIndex=32 then
  begin
  button6.Visible:=false;
  end;
end;

procedure TForm1.ComboBox6Change(Sender: TObject);
begin
button7.Visible:=true;
combobox7.Enabled:=true;
if combobox6.ItemIndex=32 then
  begin
  button7.Visible:=false;
  end;
end;

procedure TForm1.ComboBox7Change(Sender: TObject);
begin
button8.Visible:=true;
combobox8.Enabled:=true;
if combobox7.ItemIndex=32 then
  begin
  button8.Visible:=false;
  end;
end;

procedure TForm1.ComboBox8Change(Sender: TObject);
begin
button9.Visible:=true;
combobox9.Enabled:=true;
if combobox8.ItemIndex=32 then
  begin
  button9.Visible:=false;
  end;
end;

procedure TForm1.FormActivate(Sender: TObject);
var i,LCID:integer;
begin
//проверка на вшивость
version:='1.11';
form1.caption:= 'Оптимизация насадочного абсорбера ' + version;
form1.top:=10;
form1.left:=10;
LCID := GetUserDefaultLCID;
Excel.Visible[LCID]:=false;
Excel.Workbooks.Add('c:\BD\abmain.xls',LCID);
zn:=0;
zn:= Excel.Range['A1','A1'].Value2;
if zn<>1 then
begin
showmessage('IP адрес компьютера не соотвествует www.kstu.ru');
close;
Excel.Application.Quit;
end;
Excel.Application.Quit;

//

nasn:=1;
//Form2:=TForm2.Create(Self);
//Form2.ShowModal;
//form2.Hide;
//
//a[1]:=100.0;
//a[2]:=65.0;
//a[3]:=48.0;
//
a[1]:=110.0;
a[2]:=80.0;
a[3]:=60.0;
a[4]:=440.0;
a[5]:=330.0;
a[6]:=200.0;
a[7]:=140.0;
a[8]:=90.0;
a[9]:=500.0;
a[10]:=350.0;
a[11]:=220.0;
a[12]:=110.0;
a[13]:=220.0;
a[14]:=165.0;
a[15]:=120.0;
a[16]:=96.0;
a[17]:=380.0;
a[18]:=235.0;
a[19]:=170.0;
a[20]:=108.0;
a[21]:=460.0;
a[22]:=260.0;
a[23]:=165.0;
a[24]:=625.0;
a[25]:=335.0;
a[26]:=255.0;
a[27]:=195.0;
a[28]:=118.0;

a[30]:=100.0;
a[31]:=65.0;
a[32]:=48.0;
//
//eps[1]:=0.55;
//eps[2]:=0.68;
//eps[3]:=0.77;
//
eps[1]:=0.735;
eps[2]:=0.72;
eps[3]:=0.72;
eps[4]:=0.7;
eps[5]:=0.7;
eps[6]:=0.74;
eps[7]:=0.78;
eps[8]:=0.785;
eps[9]:=0.88;
eps[10]:=0.92;
eps[11]:=0.92;
eps[12]:=0.95;
eps[13]:=0.74;
eps[14]:=0.76;
eps[15]:=0.78;
eps[16]:=0.79;
eps[17]:=0.9;
eps[18]:=0.9;
eps[19]:=0.9;
eps[20]:=0.9;
eps[21]:=0.68;
eps[22]:=0.69;
eps[23]:=0.7;
eps[24]:=0.78;
eps[25]:=0.77;
eps[26]:=0.775;
eps[27]:=0.81;
eps[28]:=0.79;
//
eps[30]:=0.55;
eps[31]:=0.68;
eps[32]:=0.77;

//de[1]:=0.022;
//de[2]:=0.042;
//de[3]:=0.064;
//
de[1]:=0.027;
de[2]:=0.036;
de[3]:=0.048;
de[4]:=0.006;
de[5]:=0.009;
de[6]:=0.015;
de[7]:=0.022;
de[8]:=0.035;
de[9]:=0.007;
de[10]:=0.012;
de[11]:=0.017;
de[12]:=0.035;
de[13]:=0.014;
de[14]:=0.018;
de[15]:=0.026;
de[16]:=0.033;
de[17]:=0.01;
de[18]:=0.015;
de[19]:=0.021;
de[20]:=0.033;
de[21]:=0.006;
de[22]:=0.011;
de[23]:=0.017;
de[24]:=0.005;
de[25]:=0.009;
de[26]:=0.012;
de[27]:=0.017;
de[28]:=0.027;
de[30]:=0.022;
de[31]:=0.042;
de[32]:=0.064;

//
//nasa[1]:=0.0;
//nasa[2]:=0.0;
//nasa[3]:=0.0;
nasa[1]:=0.0;
nasa[2]:=0.0;
nasa[3]:=0.0;
nasa[4]:=-0.073;
nasa[5]:=-0.073;
nasa[6]:=-0.073;
nasa[7]:=-0.073;
nasa[8]:=-0.073;
nasa[9]:=-0.073;
nasa[10]:=-0.073;
nasa[11]:=-0.073;
nasa[12]:=-0.073;
nasa[13]:=-0.49;
nasa[14]:=-0.49;
nasa[15]:=-0.49;
nasa[16]:=-0.49;
nasa[17]:=-0.49;
nasa[18]:=-0.49;
nasa[19]:=-0.49;
nasa[20]:=-0.49;
nasa[21]:=-0.205;
nasa[22]:=-0.33;
nasa[23]:=-0.46;
nasa[24]:=-0.205;
nasa[25]:=-0.27;
nasa[26]:=-0.33;
nasa[27]:=-0.46;
nasa[28]:=-0.58;
nasa[30]:=0.0;
nasa[31]:=0.0;
nasa[32]:=0.0;
//
//nasb[1]:=1.75;
//nasb[2]:=1.75;
//nasb[3]:=1.75;
nasb[1]:=1.75;
nasb[2]:=1.75;
nasb[3]:=1.75;
nasb[4]:=1.75;
nasb[5]:=1.75;
nasb[6]:=1.75;
nasb[7]:=1.75;
nasb[8]:=1.75;
nasb[9]:=1.75;
nasb[10]:=1.75;
nasb[11]:=1.75;
nasb[12]:=1.75;
nasb[13]:=1.04;
nasb[14]:=1.04;
nasb[15]:=1.04;
nasb[16]:=1.04;
nasb[17]:=1.04;
nasb[18]:=1.04;
nasb[19]:=1.04;
nasb[20]:=1.04;
nasb[21]:=1.04;
nasb[22]:=1.04;
nasb[23]:=1.04;
nasb[24]:=1.04;
nasb[25]:=1.04;
nasb[26]:=1.04;
nasb[27]:=1.04;
nasb[28]:=1.04;
nasb[30]:=1.75;
nasb[31]:=1.75;
nasb[32]:=1.75;
//
for i:=1 to 100 do
  begin
  nasqef[i]:=0.000022;
  end;
//
standd[1]:=0.4;
standd[2]:=0.5;
standd[3]:=0.6;
standd[4]:=0.8;
standd[5]:=1.0;
standd[6]:=1.2;
standd[7]:=1.4;
standd[8]:=1.6;
standd[9]:=1.8;
standd[10]:=2.0;
standd[11]:=2.2;
standd[12]:=2.4;
standd[13]:=2.6;
standd[14]:=2.8;
standd[15]:=3.0;
standd[16]:=3.2;
standd[17]:=3.4;
standd[18]:=3.6;
standd[19]:=3.8;
standd[20]:=4.0;
standd[21]:=4.5;
standd[22]:=5.0;
standd[23]:=5.5;
standd[24]:=6.0;
standd[25]:=6.4;
standd[26]:=7.0;
standd[27]:=8.0;
standd[28]:=9.0;
//
for i:=30 to 32 do
  begin
  nasup[i]:=1;  // хордовые
  end;
for i:=1 to 3 do
  begin
  nasup[i]:=2;  // регфлярные кольца
  end;
for i:=4 to 20 do
  begin
  nasup[i]:=3;    // внавал колца
  end;
for i:=21 to 28 do
  begin
  nasup[i]:=4;      // седла
  end;

  //

nasp[1]:=0.0194;
nasp[2]:=0.0087;
nasp[3]:=0.0078;
nasp[4]:=0.0445;
nasp[5]:=0.0419;
nasp[6]:=0.0367;
nasp[7]:=0.0317;
nasp[8]:=0.024;
nasp[9]:=0.0445;
nasp[10]:=0.0419;
nasp[11]:=0.0367;
nasp[12]:=0.024;
nasp[13]:=0.021;
nasp[14]:=0.021;
nasp[15]:=0.021;
nasp[16]:=0.021;
nasp[17]:=0.021;
nasp[18]:=0.021;
nasp[19]:=0.021;
nasp[20]:=0.021;
nasp[21]:=0.021;
nasp[22]:=0.021;
nasp[23]:=0.021;
nasp[24]:=0.021;
nasp[25]:=0.021;
nasp[26]:=0.021;
nasp[27]:=0.021;
nasp[28]:=0.021;
nasp[30]:=0.0078;
nasp[31]:=0.0078;
nasp[32]:=0.0078;
//
nasq[1]:=0.0086;
nasq[2]:=0.0113;
nasq[3]:=0.0146;
nasq[4]:=0.0066;
nasq[5]:=0.0072;
nasq[6]:=0.0086;
nasq[7]:=0.01;
nasq[8]:=0.012;
nasq[9]:=0.0066;
nasq[10]:=0.0072;
nasq[11]:=0.0086;
nasq[12]:=0.012;
nasq[13]:=0.0116;
nasq[14]:=0.0116;
nasq[15]:=0.0116;
nasq[16]:=0.0116;
nasq[17]:=0.0116;
nasq[18]:=0.0116;
nasq[19]:=0.0116;
nasq[20]:=0.0116;
nasq[21]:=0.0116;
nasq[22]:=0.0116;
nasq[23]:=0.0116;
nasq[24]:=0.0116;
nasq[25]:=0.0116;
nasq[26]:=0.0116;
nasq[27]:=0.0116;
nasq[28]:=0.0116;
nasq[30]:=0.0146;
nasq[31]:=0.0146;
nasq[32]:=0.0146;
  //
nasbn[1]:=173;
nasbn[2]:=144;
nasbn[3]:=119;
nasbn[4]:=193;
nasbn[5]:=190;
nasbn[6]:=184;
nasbn[7]:=178;
nasbn[8]:=196;
nasbn[9]:=193;
nasbn[10]:=190;
nasbn[11]:=184;
nasbn[12]:=169;
nasbn[13]:=126;
nasbn[14]:=126;
nasbn[15]:=126;
nasbn[16]:=126;
nasbn[17]:=126;
nasbn[18]:=126;
nasbn[19]:=126;
nasbn[20]:=126;
nasbn[21]:=30;
nasbn[22]:=30;
nasbn[23]:=30;
nasbn[24]:=35.5;
nasbn[25]:=34.2;
nasbn[26]:=33;
nasbn[27]:=30.4;
nasbn[28]:=28;
nasbn[30]:=119;
nasbn[31]:=119;
nasbn[32]:=119;


//
he[1]:=0.05;
he[2]:=0.08;
he[3]:=0.1;
he[4]:=0.010;
he[5]:=0.015;
he[6]:=0.025;
he[7]:=0.035;
he[8]:=0.050;
he[9]:=0.010;
he[10]:=0.015;
he[11]:=0.025;
he[12]:=0.050;
he[13]:=0.025;
he[14]:=0.035;
he[15]:=0.050;
he[16]:=0.060;
he[17]:=0.015;
he[18]:=0.025;
he[19]:=0.035;
he[20]:=0.05;
he[21]:=0.0125;
he[22]:=0.025;
he[23]:=0.038;
he[24]:=0.0125;
he[25]:=0.019;
he[26]:=0.025;
he[27]:=0.038;
he[28]:=0.050;
he[30]:=0.1;
he[31]:=0.1;
he[32]:=0.1;

end;
//

//
procedure TForm1.ComboBox9Change(Sender: TObject);
begin
button10.Visible:=true;
if combobox9.ItemIndex=32 then
  begin
  button10.Visible:=false;
  end;
  //

end;

procedure TForm1.Button4Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox1.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox1.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox1.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox1.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox1.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox1.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox1.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox1.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox1.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox1.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox1.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.he[unit1.form1.combobox1.itemindex+1]);
if (unit1.Form1.ComboBox1.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[51]);
  form2.Edit2.Text:=floattostr(unit1.eps[51]);
  form2.Edit3.Text:=floattostr(unit1.de[51]);
  form2.Edit4.Text:=floattostr(unit1.nasa[51]);
  form2.Edit5.Text:=floattostr(unit1.nasb[51]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[51]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[51]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[51]);
  form2.edit8.Text:=floattostr(unit1.nasq[51]);
  form2.edit9.Text:=floattostr(unit1.nasbn[51]);
  form2.edit10.Text:=floattostr(unit1.he[51]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox1.ItemIndex=28) then
  begin
  unit1.a[51]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[51]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[51]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[51]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[51]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[51]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[51]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[51]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[51]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[51]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[51]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[51]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox2.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox2.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox2.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox2.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox2.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox2.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox2.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox2.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox2.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox2.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox2.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.he[unit1.form1.combobox2.itemindex+1]);
if (unit1.Form1.ComboBox2.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[52]);
  form2.Edit2.Text:=floattostr(unit1.eps[52]);
  form2.Edit3.Text:=floattostr(unit1.de[52]);
  form2.Edit4.Text:=floattostr(unit1.nasa[52]);
  form2.Edit5.Text:=floattostr(unit1.nasb[52]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[52]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[52]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[52]);
  form2.edit8.Text:=floattostr(unit1.nasp[52]);
  form2.edit9.Text:=floattostr(unit1.nasbn[52]);
  form2.edit10.Text:=floattostr(unit1.he[52]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox2.ItemIndex=28) then
  begin
  unit1.a[52]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[52]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[52]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[52]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[52]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[52]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[52]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[52]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[52]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[52]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[52]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[52]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox3.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox3.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox3.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox3.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox3.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox3.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox3.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox3.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox3.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox3.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox3.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.he[unit1.form1.combobox3.itemindex+1]);
if (unit1.Form1.ComboBox3.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[53]);
  form2.Edit2.Text:=floattostr(unit1.eps[53]);
  form2.Edit3.Text:=floattostr(unit1.de[53]);
  form2.Edit4.Text:=floattostr(unit1.nasa[53]);
  form2.Edit5.Text:=floattostr(unit1.nasb[53]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[53]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[53]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[53]);
  form2.edit8.Text:=floattostr(unit1.nasp[53]);
  form2.edit9.Text:=floattostr(unit1.nasbn[53]);
  form2.edit10.Text:=floattostr(unit1.he[53]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox3.ItemIndex=28) then
  begin
  unit1.a[53]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[53]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[53]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[53]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[53]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[53]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[53]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[53]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[53]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[53]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[53]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[53]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox4.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox4.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox4.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox4.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox4.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox4.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox4.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox4.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox4.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox4.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox4.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.he[unit1.form1.combobox4.itemindex+1]);
if (unit1.Form1.ComboBox4.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[54]);
  form2.Edit2.Text:=floattostr(unit1.eps[54]);
  form2.Edit3.Text:=floattostr(unit1.de[54]);
  form2.Edit4.Text:=floattostr(unit1.nasa[54]);
  form2.Edit5.Text:=floattostr(unit1.nasb[54]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[54]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[54]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[54]);
  form2.edit8.Text:=floattostr(unit1.nasp[54]);
  form2.edit9.Text:=floattostr(unit1.nasbn[54]);
  form2.edit10.Text:=floattostr(unit1.he[54]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox4.ItemIndex=28) then
  begin
  unit1.a[54]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[54]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[54]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[54]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[54]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[54]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[54]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[54]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[54]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[54]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[54]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[54]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox5.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox5.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox5.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox5.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox5.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox5.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox5.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox5.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox5.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox5.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox5.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.He[unit1.form1.combobox5.itemindex+1]);
if (unit1.Form1.ComboBox5.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[55]);
  form2.Edit2.Text:=floattostr(unit1.eps[55]);
  form2.Edit3.Text:=floattostr(unit1.de[55]);
  form2.Edit4.Text:=floattostr(unit1.nasa[55]);
  form2.Edit5.Text:=floattostr(unit1.nasb[55]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[55]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[55]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[55]);
  form2.edit8.Text:=floattostr(unit1.nasp[55]);
  form2.edit9.Text:=floattostr(unit1.nasbn[55]);
  form2.edit10.Text:=floattostr(unit1.he[55]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox5.ItemIndex=28) then
  begin
  unit1.a[55]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[55]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[55]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[55]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[55]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[55]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[55]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[55]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[55]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[55]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[55]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[55]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox6.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox6.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox6.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox6.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox6.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox6.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox6.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox6.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox6.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox6.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox6.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.He[unit1.form1.combobox6.itemindex+1]);
if (unit1.Form1.ComboBox6.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[56]);
  form2.Edit2.Text:=floattostr(unit1.eps[56]);
  form2.Edit3.Text:=floattostr(unit1.de[56]);
  form2.Edit4.Text:=floattostr(unit1.nasa[56]);
  form2.Edit5.Text:=floattostr(unit1.nasb[56]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[56]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[56]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[56]);
  form2.edit8.Text:=floattostr(unit1.nasp[56]);
  form2.edit9.Text:=floattostr(unit1.nasbn[56]);
  form2.edit10.Text:=floattostr(unit1.he[56]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox6.ItemIndex=28) then
  begin
  unit1.a[56]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[56]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[56]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[56]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[56]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[56]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[56]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[56]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[56]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[56]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[56]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[56]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
//form2.Close;
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox7.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox7.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox7.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox7.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox7.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox7.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox7.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox7.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox7.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox7.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox7.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.He[unit1.form1.combobox7.itemindex+1]);
if (unit1.Form1.ComboBox7.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[57]);
  form2.Edit2.Text:=floattostr(unit1.eps[57]);
  form2.Edit3.Text:=floattostr(unit1.de[57]);
  form2.Edit4.Text:=floattostr(unit1.nasa[57]);
  form2.Edit5.Text:=floattostr(unit1.nasb[57]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[57]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[57]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[57]);
  form2.edit8.Text:=floattostr(unit1.nasp[57]);
  form2.edit9.Text:=floattostr(unit1.nasbn[57]);
  form2.edit10.Text:=floattostr(unit1.he[57]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox7.ItemIndex=28) then
  begin
  unit1.a[57]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[57]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[57]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[57]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[57]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[57]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[57]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[57]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[57]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[57]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[57]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[57]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox8.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox8.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox8.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox8.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox8.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox8.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox8.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox8.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox8.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox8.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox8.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.He[unit1.form1.combobox8.itemindex+1]);
if (unit1.Form1.ComboBox8.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[58]);
  form2.Edit2.Text:=floattostr(unit1.eps[58]);
  form2.Edit3.Text:=floattostr(unit1.de[58]);
  form2.Edit4.Text:=floattostr(unit1.nasa[58]);
  form2.Edit5.Text:=floattostr(unit1.nasb[58]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[58]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[58]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[58]);
  form2.edit8.Text:=floattostr(unit1.nasp[58]);
  form2.edit9.Text:=floattostr(unit1.nasbn[58]);
  form2.edit10.Text:=floattostr(unit1.he[58]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox8.ItemIndex=28) then
  begin
  unit1.a[58]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[58]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[58]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[58]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[58]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[58]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[58]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[58]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[58]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[58]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[58]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[58]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
form2.Edit1.Enabled:=false;
form2.Edit2.Enabled:=false;
form2.Edit3.Enabled:=false;
form2.Edit4.Enabled:=false;
form2.Edit5.Enabled:=false;
form2.Edit6.Enabled:=false;
unit2.form2.Edit7.Enabled:=false;
form2.edit8.Enabled:=false;
form2.edit9.Enabled:=false;
form2.edit10.Enabled:=false;
form2.combobox1.Enabled:=false;
unit2.Form2.Button1.Enabled:=false;
form2.Label1.Caption:=form1.ComboBox9.Text;
form2.Edit1.Text:=floattostr(unit1.a[form1.ComboBox9.ItemIndex+1]);
form2.Edit2.Text:=floattostr(unit1.eps[form1.combobox9.ItemIndex+1]);
form2.Edit3.Text:=floattostr(unit1.de[form1.ComboBox9.ItemIndex+1]);
form2.Edit4.Text:=floattostr(unit1.nasa[form1.ComboBox9.ItemIndex+1]);
form2.Edit5.Text:=floattostr(unit1.nasb[form1.ComboBox9.ItemIndex+1]);
form2.Edit6.Text:=floattostr(unit1.nasqef[form1.ComboBox9.ItemIndex+1]);
unit2.Form2.combobox1.ItemIndex:=nasup[unit1.Form1.ComboBox9.ItemIndex+1]-1;
form2.Edit7.text:=floattostr(unit1.nasp[unit1.form1.combobox9.itemindex+1]);
form2.Edit8.text:=floattostr(unit1.nasq[unit1.form1.combobox9.itemindex+1]);
form2.Edit9.text:=floattostr(unit1.nasbn[unit1.form1.combobox9.itemindex+1]);
form2.Edit10.text:=floattostr(unit1.He[unit1.form1.combobox9.itemindex+1]);
if (unit1.Form1.ComboBox9.ItemIndex=28) then
  begin
  form2.Label1.Caption:='Параметры введены вручную';
  form2.Edit1.Text:=floattostr(unit1.a[59]);
  form2.Edit2.Text:=floattostr(unit1.eps[59]);
  form2.Edit3.Text:=floattostr(unit1.de[59]);
  form2.Edit4.Text:=floattostr(unit1.nasa[59]);
  form2.Edit5.Text:=floattostr(unit1.nasb[59]);
  form2.Edit6.Text:=floattostr(unit1.nasqef[59]);
  unit2.Form2.ComboBox1.ItemIndex:=unit1.nasup[59]-1;
  form2.edit7.Text:=floattostr(unit1.nasp[59]);
  form2.edit8.Text:=floattostr(unit1.nasp[59]);
  form2.edit9.Text:=floattostr(unit1.nasbn[59]);
  form2.edit10.Text:=floattostr(unit1.he[59]);
  form2.Edit1.Enabled:=true;
  form2.Edit2.Enabled:=true;
  form2.Edit3.Enabled:=true;
  form2.Edit4.Enabled:=true;
  form2.Edit5.Enabled:=true;
  form2.Edit6.Enabled:=true;
  form2.Edit7.Enabled:=true;
  form2.Edit8.Enabled:=true;
  form2.edit9.Enabled:=true;
  form2.edit10.Enabled:=true;
  unit2.Form2.Button1.Enabled:=true;
  unit2.Form2.combobox1.Enabled:=true;
  end;
Form2.ShowModal; // (или Form2.ShowModal) показ Формы
if (unit1.Form1.ComboBox9.ItemIndex=28) then
  begin
  unit1.a[59]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.a[59]:=strtofloat(unit2.Form2.Edit1.Text);
  unit1.eps[59]:=strtofloat(unit2.Form2.Edit2.Text);
  unit1.de[59]:=strtofloat(unit2.Form2.Edit3.Text);
  unit1.nasa[59]:=strtofloat(unit2.Form2.Edit4.Text);
  unit1.nasb[59]:=strtofloat(unit2.Form2.Edit5.Text);
  unit1.nasqef[59]:=strtofloat(unit2.Form2.Edit6.Text);
  unit1.nasup[59]:=unit2.Form2.ComboBox1.ItemIndex+1;
  unit1.nasp[59]:=strtofloat(unit2.form2.edit7.text);
  unit1.nasq[59]:=strtofloat(unit2.form2.edit8.text);
  unit1.nasbn[59]:=strtofloat(unit2.form2.edit9.text);
  unit1.he[59]:=strtofloat(unit2.form2.edit10.text);
  end;
end;

procedure TForm1.Edit1Change(Sender: TObject);
var i,LCID:integer;
begin
//проверка на вшивость
if zn<>1 then {количество веществ в базе данных}
begin
showmessage('проверка на вшивость не пройдена');
close;
Excel.Application.Quit;
end;
Excel.Application.Quit;
//
end;

end.
