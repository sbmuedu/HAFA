<?php
/**
 * تنظیمات افزوده حافا برای MediaWiki 1.46
 * این فایل را به انتهای LocalSettings.php فعلی اضافه کنید؛
 * مقادیر رازدار (گذرواژه DB، SecretKey و UpgradeKey) در این فایل وجود ندارند.
 */

# پیش از انتشار، نام میزبان واقعی و HTTPS را جایگزین کنید.
# $wgServer = 'https://wiki.example.edu';
$wgLocaltimezone = 'Asia/Tehran';

# بارگذاری فایل‌های مستندات حافا
$wgEnableUploads = true;
$wgFileExtensions = array_values( array_unique( array_merge(
    $wgFileExtensions,
    [ 'pdf', 'png', 'jpg', 'jpeg', 'svg', 'bpmn' ]
) ) );
$wgMaxUploadSize = 20 * 1024 * 1024; // 20 MiB

# فضاهای نام حافا. قبل از استفاده، از آزاد بودن شناسه‌ها در تنظیمات فعلی اطمینان یابید.
define( 'NS_HAFA_PROCESS', 3080 );
define( 'NS_HAFA_ACTIVITY', 3082 );
define( 'NS_HAFA_SYSTEM', 3084 );
define( 'NS_HAFA_TERM', 3086 );
define( 'NS_HAFA_DATA', 3088 );
define( 'NS_HAFA_NEED', 3090 );
define( 'NS_HAFA', 3092 );
define( 'NS_HAFA_CAPABILITY', 3094 );
define( 'NS_HAFA_GAP', 3096 );

$wgExtraNamespaces[NS_HAFA_PROCESS] = 'فرآیند';
$wgExtraNamespaces[NS_HAFA_PROCESS + 1] = 'بحث_فرآیند';
$wgExtraNamespaces[NS_HAFA_ACTIVITY] = 'فعالیت';
$wgExtraNamespaces[NS_HAFA_ACTIVITY + 1] = 'بحث_فعالیت';
$wgExtraNamespaces[NS_HAFA_SYSTEM] = 'سامانه';
$wgExtraNamespaces[NS_HAFA_SYSTEM + 1] = 'بحث_سامانه';
$wgExtraNamespaces[NS_HAFA_TERM] = 'واژه';
$wgExtraNamespaces[NS_HAFA_TERM + 1] = 'بحث_واژه';
$wgExtraNamespaces[NS_HAFA_DATA] = 'داده';
$wgExtraNamespaces[NS_HAFA_DATA + 1] = 'بحث_داده';
$wgExtraNamespaces[NS_HAFA_NEED] = 'نیاز';
$wgExtraNamespaces[NS_HAFA_NEED + 1] = 'بحث_نیاز';
$wgExtraNamespaces[NS_HAFA] = 'حافا';
$wgExtraNamespaces[NS_HAFA + 1] = 'بحث_حافا';
$wgExtraNamespaces[NS_HAFA_CAPABILITY] = 'قابلیت';
$wgExtraNamespaces[NS_HAFA_CAPABILITY + 1] = 'بحث_قابلیت';
$wgExtraNamespaces[NS_HAFA_GAP] = 'شکاف';
$wgExtraNamespaces[NS_HAFA_GAP + 1] = 'بحث_شکاف';

foreach ( [ NS_HAFA_PROCESS, NS_HAFA_ACTIVITY, NS_HAFA_SYSTEM, NS_HAFA_TERM, NS_HAFA_DATA, NS_HAFA_NEED, NS_HAFA, NS_HAFA_CAPABILITY, NS_HAFA_GAP ] as $hafaNamespace ) {
    $wgNamespacesWithSubpages[$hafaNamespace] = true;
}

# گروه‌های نقش حافا. جزئیات نهایی دسترسی در صفحه «حافا: معماری ابزار و سطوح دسترسی» تصویب شود.
$wgGroupPermissions['hafa-contributor']['read'] = true;
$wgGroupPermissions['hafa-contributor']['edit'] = true;
$wgGroupPermissions['hafa-contributor']['createpage'] = true;
$wgGroupPermissions['hafa-contributor']['upload'] = true;
$wgGroupPermissions['hafa-unit-representative'] = $wgGroupPermissions['hafa-contributor'];
$wgGroupPermissions['hafa-reviewer'] = $wgGroupPermissions['hafa-contributor'];
$wgGroupPermissions['hafa-wiki-steward'] = $wgGroupPermissions['hafa-reviewer'];
$wgGroupPermissions['hafa-wiki-steward']['editinterface'] = true;
$wgGroupPermissions['hafa-wiki-steward']['move'] = true;

# حساب اختصاصی ورود گروهی؛ از sysop استفاده نشود.
$wgGroupPermissions['hafa-importer']['read'] = true;
$wgGroupPermissions['hafa-importer']['edit'] = true;
$wgGroupPermissions['hafa-importer']['createpage'] = true;
$wgGroupPermissions['hafa-importer']['import'] = true;
$wgGroupPermissions['hafa-importer']['importupload'] = true;
$wgGroupPermissions['hafa-importer']['upload'] = true;

# مدیر دسترسی حافا: مدیریت عضویت کاربران فقط در گروه‌های عملیاتی حافا.
$wgGroupPermissions['hafa-access-manager']['read'] = true;
$wgGroupPermissions['hafa-access-manager']['userrights'] = true;

$hafaAssignableGroups = [
    'hafa-contributor',
    'hafa-unit-representative',
    'hafa-reviewer',
    'hafa-wiki-steward',
    'hafa-importer',
];
$wgAddGroups['hafa-access-manager'] = $hafaAssignableGroups;
$wgRemoveGroups['hafa-access-manager'] = $hafaAssignableGroups;

# راه‌اندازی اولیه: این حق را فقط تا انتساب نخستین مدیر دسترسی حافا به یک sysop بدهید.
# سپس این خط را دوباره کامنت کنید تا همه sysopها نتوانند دسترسی کاربران را تغییر دهند.
# $wgGroupPermissions['sysop']['userrights'] = true;
# $wgAddGroups['sysop'] = array_merge( $wgAddGroups['sysop'] ?? [], array_merge( $hafaAssignableGroups, [ 'hafa-access-manager' ] ) );
# $wgRemoveGroups['sysop'] = array_merge( $wgRemoveGroups['sysop'] ?? [], array_merge( $hafaAssignableGroups, [ 'hafa-access-manager' ] ) );

# اگر حساب شما عضو bureaucrat است، این گروه می‌تواند گروه‌های حافا را مدیریت کند.
$wgAddGroups['bureaucrat'] = array_merge( $wgAddGroups['bureaucrat'] ?? [], array_merge( $hafaAssignableGroups, [ 'hafa-access-manager' ] ) );
$wgRemoveGroups['bureaucrat'] = array_merge( $wgRemoveGroups['bureaucrat'] ?? [], array_merge( $hafaAssignableGroups, [ 'hafa-access-manager' ] ) );

# حساب فقط توسط مدیر ویکی ساخته می‌شود؛ حساب مشترک ممنوع است.
$wgGroupPermissions['*']['createaccount'] = false;
$wgGroupPermissions['sysop']['createaccount'] = true;

# محدودسازی صفحات ساختاری؛ فقط مدیران و راهبر ویکی مجاز به ویرایش‌اند.
$wgGroupPermissions['hafa-wiki-steward']['editprotected'] = true;
$wgGroupPermissions['sysop']['editprotected'] = true;

# افزونه‌های اختیاری: فقط بعد از نصب واقعی و اجرای update.php، خطوط مربوط را فعال کنید.
# wfLoadExtension( 'ApprovedRevs' );
# wfLoadExtension( 'ConfirmEdit' );
# wfLoadExtension( 'AbuseFilter' );
# wfLoadExtension( 'OATHAuth' );

# توجه: مقدار قبلی $smwgEditProtectionRight = false را حذف کنید،
# مگر آنکه نیاز و اثر آن در نسخه نصب‌شده Semantic MediaWiki آزموده شده باشد.
