############################################################
# リマインダーツール
############################################################
# === 各種設定値 ==========
# 指定時刻のどのくらい前からリマインドするか
$REMINDER_MINUTES = 30;

$STOP_SECONDS = 3;
$TARGET_FOLDER_PATH = "D:\workspace_ps\リマインダーツール";
$TARGET_FILE_PATH = "reminder_list.txt";


# ======================================
# リマインダーリスト取得
# ======================================
function GetReminderList{
    $file = (Get-Content -Encoding utf8 $TARGET_FOLDER_PATH\$TARGET_FILE_PATH) -as [string[]];
    $reminderHash = @{};
    
    foreach($line in $file) {
        $arr = $line.split(" ");
        
        if([System.String]::IsNullOrWhiteSpace($arr[0])) {
            continue;
        }
        
        $reminderHash.add($arr[0], $line);
    }
    
    return $reminderHash;
}

# ======================================
# リマインダーリストの年月日時分秒取得
#   指定時刻n分前を設定
# ======================================
function GetReminderDateTime{
    param($key)
    
    $keyArr = $key.split(":");
    $keyHour = $keyArr[0].PadLeft(2, "0");
    $keyMinute = $keyArr[1].PadLeft(2, "0");
    
    $nowMonth = ([string](Get-Date).Month).PadLeft(2, "0");
    $nowDay = ([string](Get-Date).Day).PadLeft(2, "0");
    
    $reminderDateTimeStr = "{0}{1}{2}{3}{4}{5}" -f (Get-Date).Year, $nowMonth, $nowDay, $keyHour, $keyMinute, "00";
    $reminderDateTime = [DateTime]::ParseExact($reminderDateTimeStr, "yyyyMMddHHmmss", $null);
    
    return $reminderDateTime;
}

# ======================================
# リマインダー表示
# ======================================
function PopupReminder{

    while($true){
        
        # 現在時刻取得 (秒は常に00固定)
        $nowDateTime = [DateTime]::ParseExact((Get-Date).ToString("yyyyMMddHHmm") + "00", "yyyyMMddHHmmss", $null);
        Write-Host "現在時刻" $nowDateTime;
        
        foreach($key in $reminderHash.Keys){
            # リマインダーリストの年月日時分秒
            $reminderDateTime = GetReminderDateTime $key;
            Write-Host "リマインダー" $reminderDateTime;
            
            # 現在日時がリマインダー時刻以降の場合
            $diffTime = $reminderDateTime - $nowDateTime;
            $diffMinutes = [Math]::Abs($diffTime.Minutes);
            
            if($diffMinutes -le $REMINDER_MINUTES){
                $val = $reminderHash[$key];
                Write-Host "差異" $diffMinutes;
                Write-Host $val;
                
                # メッセージ表示
                $wsobj = new-object -comobject wscript.shell;
                $result = $wsobj.popup($val, 2, "リマインダー");
            }
        }
        
        #イベント実行単位
        Start-Sleep -Seconds $STOP_SECONDS;
    }
}


# === リマインダー実行判定 ==========
[string]$res = Read-Host "リマインダーツールを実行しますか？ (y:実行する, 左記以外:実行しない)";

if($res -ne "y"){
     Write-Host "リマインダーツールを実行しない";
     Exit;
}

# === リマインダーツール実行 ==========
# リマインダーリストファイル読み込み
Write-Host "本日のリマインダーリスト";
$reminderHash = GetReminderList;

# イベント発生待ち
PopupReminder;


Read-Host "終了するには何かキーを押してください";
exit;

