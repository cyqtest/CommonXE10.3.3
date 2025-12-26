unit TU2.Helper.Currency;

interface

function CurrencyToChineseCapitalCharacter(const AValue: Currency; const ADecimals: Cardinal=4): string;
function CurrencyToString(const AValue: Currency; const ADecimals: Cardinal=4): string;

implementation

uses System.SysUtils, System.Math;

function CurrencyRound(var U: UInt64; const ADecimals: Cardinal): Integer; inline;
var
  W: UInt64;
begin//Bankers-rounding
  Result := 4-ADecimals;
  if Result<0 then
    Result := 0
  else if Result>0 then
  begin
    case Result of
      1:begin   //li
        DivMod(U, 10, U, W);
        if (W > 5) or ((W = 5) and Odd(U)) then
          Inc(U);
      end;
      2:begin  //fen
        DivMod(U, 100, U, W);
        if (W > 50) or ((W = 50) and Odd(U)) then
          Inc(U);
      end;
      3:begin  //jiao
        DivMod(U, 1000, U, W);
        if (W > 500) or ((W = 500) and Odd(U)) then
          Inc(U);
      end;
      4:begin  //yuan
        DivMod(U, 10000, U, W);
        if (W > 5000) or ((W = 5000) and Odd(U)) then
          Inc(U);
      end;
    end;
  end;
end;

function CurrencyToChineseCapitalCharacter(const AValue: Currency; const ADecimals: Cardinal=4): string;
const//Currency: [-922337203685477.5807, 922337203685477.5807]
  CCCNegative = '¸º';
  CCCZheng = 'Õû';
  CCCNumbers: array[0..9] of Char = ('Áã','Ò¼','·¡','Èþ','ËÁ','Îé','Â½','Æâ','°Æ','¾Á');
  CCCUnits: array[0..18] of Char = ('ºÁ', 'Àå', '·Ö', '½Ç', 'Ô²','Ê°','°Û','Çª','Íò',
                                     'Ê°','°Û','Çª','ÒÚ','Ê°','°Û','Çª','Íò','Õ×','Ê°');
var
  U, W: UInt64;
  Digits, Idx, ZeroFlag: Integer;
  Negative: Boolean;
  Buff: array[0..38] of Char;
begin
  U := PUInt64(@AValue)^;
  if U <> 0 then
  begin
    Negative := (U and $8000000000000000) <> 0;
    if Negative then
      U := not U + 1;
    Digits := CurrencyRound(U, ADecimals);
    if U<>0 then
    begin
      //Try skip trailing zero
      repeat
        DivMod(U, 10, U, W);
        Inc(Digits);
      until W<>0;
      Dec(Digits);
      Idx := 38;
      if Digits>=3 then
      begin
        Buff[Idx] := CCCZheng;
        Dec(Idx);
        if Digits>4 then
        begin
          Buff[Idx] := CCCUnits[4];
          Dec(Idx);
          if Digits>17 then
          begin
            Buff[Idx] := CCCUnits[17];
            Dec(Idx);
          end else if Digits>12 then
          begin
            Buff[Idx] := CCCUnits[12];
            Dec(Idx);
          end else if Digits>8 then
          begin
            Buff[Idx] := CCCUnits[8];
            Dec(Idx);
          end;
        end;
      end;
      Buff[Idx] := CCCUnits[Digits];
      Dec(Idx);
      Buff[Idx] := CCCNumbers[W];
      Dec(Idx);
      //Do Split
      ZeroFlag := 0;
      while U<>0 do
      begin
        Inc(Digits);
        DivMod(U, 10, U, W);
        if Digits in [4,8,12,17] then
        begin
          if ZeroFlag>0 then
          begin
            Buff[Idx] := CCCNumbers[0];
            Dec(Idx);
          end else if (ZeroFlag<0) and (Digits>8) then
            Inc(Idx);
          Buff[Idx] := CCCUnits[Digits];
          Dec(Idx);
          if W<>0 then
          begin
            Buff[Idx] := CCCNumbers[W];
            Dec(Idx);
            ZeroFlag := 0;
          end else
            ZeroFlag := -1;
        end else begin
          if W<>0 then
          begin
            if ZeroFlag>0 then
            begin
              Buff[Idx] := CCCNumbers[0];
              Dec(Idx);
            end;
            Buff[Idx] := CCCUnits[Digits];
            Dec(Idx);
            Buff[Idx] := CCCNumbers[W];
            Dec(Idx);
            ZeroFlag := 0;
          end else begin
            if ZeroFlag=0 then
              ZeroFlag := 1;
          end;
        end;
      end;

      if Negative then
        Buff[Idx] := CCCNegative
      else Inc(Idx);

      //Copy Result
      Digits := 38+1-idx;
      SetLength(Result, Digits);
      Move(Buff[idx], PChar(Result)^, Digits * SizeOf(WideChar));
      Exit;
    end;
  end;
  Result := CCCNumbers[0]+CCCUnits[4]+CCCZheng;
end;

function CurrencyToString(const AValue: Currency; const ADecimals: Cardinal=4): string;
const
  NegativeChar = '-';
  DecimalDotChar = '.';
var
  U: UInt64;
  Digits: Integer;
  Negative: Boolean;
begin
  U := PUInt64(@AValue)^;
  Negative := (U and $8000000000000000) <> 0;
  if Negative then
    U := not U + 1;
  Digits := CurrencyRound(U, ADecimals);
  Result := UIntToStr(U);
  if Digits<4 then
    Result := Result.Insert(Result.Length+Digits-4, DecimalDotChar);
  if Negative then
    Result := NegativeChar + Result;
end;

end.
