import 'package:flutter/gestures.dart';
import 'package:url_launcher/url_launcher.dart';
import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';
import 'package:intl/intl.dart';
import 'package:flutter/foundation.dart';
import 'package:whats_for_dino/2025Sem_1.dart';
import 'classes.dart';
import '2023Sem_2.dart';
import 'package:url_strategy/url_strategy.dart';

// Constants
final START_DATE = DateTime.utc(2025, 2, 10); // Start date of current term
final MID_DATE   = DateTime.utc(2025, 5, 25); // End of first menu, 1 day before start of new menu
final FINAL_DATE = DateTime.utc(2025, 9, 7); // Date end of second menu

final Term FIRST_TERM = term_1_2025;
final Term SECOND_TERM = term_1_2025;

final int FIRST_TERM_START_WEEK = 1;
final int SECOND_TERM_START_WEEK = 1;

// Settings Variables
bool loading = true;
bool setting_veg = false;
bool setting_dairy_allergy = false;
bool setting_gluten_allergy = false;

bool setting_dark_mode = false;
bool setting_hide_dietary_tags = false;

// Print the date nicely
Format_Date() {
  DateTime date_time_today = DateTime.now();
  DateTime date_today = DateTime.utc(
      date_time_today.year, date_time_today.month, date_time_today.day);
  if (date_today.difference(DateTextController).inDays == 0) {
    String formatted_date = dateFormatter.format(DateTextController);
    DateTextControllerFormatted = "$formatted_date (Today)";
  } else {
    DateTextControllerFormatted = dateFormatter.format(DateTextController);
  }
}

Load_settings() async {
  final prefs = await SharedPreferences.getInstance();
  setting_veg = prefs.getBool('Vegetarian') ?? false;
  setting_dairy_allergy = prefs.getBool('Dairy_Allergy') ?? false;
  setting_gluten_allergy = prefs.getBool('Gluten_Allergy') ?? false;
  setting_dark_mode = prefs.getBool('Dark_Mode') ?? false;
  setting_hide_dietary_tags = prefs.getBool('Hide_Dietary_Tags') ?? false;
}

Update_settings() async {
  final prefs = await SharedPreferences.getInstance();
  await prefs.setBool('Vegetarian', setting_veg);
  await prefs.setBool('Dairy_Allergy', setting_dairy_allergy);
  await prefs.setBool('Gluten_Allergy', setting_gluten_allergy);
  await prefs.setBool('Dark_Mode', setting_dark_mode);
  await prefs.setBool('Hide_Dietary_Tags', setting_hide_dietary_tags);
}

final dateFormatter = DateFormat('dd/MM/yyyy');

DateTime date_time_now = DateTime.now();

DateTime DateTextController =
    DateTime.utc(date_time_now.year, date_time_now.month, date_time_now.day);

var DateTextControllerFormatted =
    dateFormatter.format(DateTextController) + " (Today)";

PageController page_controller = PageController(
    initialPage: DateTextController.difference(START_DATE).inDays);
int old_page = DateTextController.difference(START_DATE).inDays;

// Updates the menu after the app is awoken from sleep incase the date has changed
class LifecycleEventHandler extends WidgetsBindingObserver {
  final AsyncCallback? resumeCallBack;
  final AsyncCallback? suspendingCallBack;

  LifecycleEventHandler({
    this.resumeCallBack,
    this.suspendingCallBack,
  });

  @override
  Future<void> didChangeAppLifecycleState(AppLifecycleState state) async {
    switch (state) {
      case AppLifecycleState.resumed:
        if (resumeCallBack != null) {
          await resumeCallBack!();
        }
        break;
      case AppLifecycleState.inactive:
      case AppLifecycleState.paused:
      case AppLifecycleState.detached:
        if (suspendingCallBack != null) {
          await suspendingCallBack!();
        }
        break;
    }
  }
}

void main() async {
  setPathUrlStrategy();
  WidgetsFlutterBinding.ensureInitialized();
  await Load_settings();
  runApp(const MyApp());
}

class MyApp extends StatefulWidget {
  const MyApp({super.key});

  @override
  State<MyApp> createState() => _MyAppState();
}

class _MyAppState extends State<MyApp> with WidgetsBindingObserver {
  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addObserver(LifecycleEventHandler(
        resumeCallBack: () async => setState(() {
              Format_Date();
            })));
  }

  @override
  void dispose() {
    WidgetsBinding.instance.removeObserver(this);
    super.dispose();
  }

  void Update_App_State() {
    setState(() {});
  }

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      initialRoute: '/',
      onGenerateRoute: (RouteSettings settings) {
        switch (settings.name) {
          case '/':
            return MaterialPageRoute(
                settings: settings,
                builder: (context) => MyHomePage(
                    title: 'What\'s for Dino',
                    ChangeDarkMode: Update_App_State));
          case '/settings':
            final Function ChangeDarkMode = settings.arguments as Function;
            return MaterialPageRoute(
                settings: settings,
                builder: (context) =>
                    SettingsPage(ChangeDarkMode: ChangeDarkMode));
          default:
            return MaterialPageRoute(
                settings: settings,
                builder: (context) => MyHomePage(
                    title: 'What\'s for Dino',
                    ChangeDarkMode: Update_App_State));
        }
      },
      title: 'What\'s for Dino',
      theme: ThemeData(
        useMaterial3: false,
        primarySwatch: Colors.green,
        brightness:
            setting_dark_mode == true ? Brightness.dark : Brightness.light,
      ),
    );
  }
}

class MyHomePage extends StatefulWidget {
  const MyHomePage(
      {super.key, required this.title, required this.ChangeDarkMode});

  final String title;
  final Function ChangeDarkMode;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  @override
  void initState() {
    super.initState();
    if (DateTextController.isAfter(FINAL_DATE) == true) {
      DateTextController = FINAL_DATE;
      Format_Date();
      old_page = FINAL_DATE.difference(START_DATE).inDays;
      WidgetsBinding.instance.addPostFrameCallback(
        (_) {
          showDialog(
            context: context,
            builder: (BuildContext context) {
              return AlertDialog(
                title: Center(child: Text("WARNING")),
                content: Text(
                    textAlign: TextAlign.center,
                    style: Theme.of(context)
                        .textTheme
                        .titleMedium!
                        .copyWith(fontWeight: FontWeight.w600, fontSize: 14),
                    "The current date isn't available on the menu, seems you like you may need to update the app!"),
                actions: [
                  Center(
                    child: TextButton(
                      child: Text("OK"),
                      onPressed: () {
                        Navigator.pop(context);
                      },
                    ),
                  ),
                ],
              );
            },
          );
        },
      );
    } else if (DateTextController.isBefore(START_DATE) == true) {
      DateTextController = START_DATE;
      Format_Date();
      old_page = 0;
      WidgetsBinding.instance.addPostFrameCallback(
        (_) {
          showDialog(
            context: context,
            builder: (BuildContext context) {
              return AlertDialog(
                title: Center(child: Text("WARNING")),
                content: Text(
                    textAlign: TextAlign.center,
                    style: Theme.of(context)
                        .textTheme
                        .titleMedium!
                        .copyWith(fontWeight: FontWeight.w600, fontSize: 14),
                    "The current date isn't available on the menu, seems like you have been messing around with the date settings!"),
                actions: [
                  Center(
                    child: TextButton(
                      child: Text("OK"),
                      onPressed: () {
                        Navigator.pop(context);
                      },
                    ),
                  ),
                ],
              );
            },
          );
        },
      );
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Center(child: Text(widget.title)),
      ),
      body: Column(
        children: [
          Center(
            child: Row(
              children: [
                Expanded(
                  flex: 1,
                  child: Align(
                    alignment: Alignment.centerLeft,
                    child: IconButton(
                      icon: const Icon(Icons.settings),
                      onPressed: () {
                        Navigator.pushNamed(
                          context,
                          '/settings',
                          arguments: widget.ChangeDarkMode,
                        ).then((value) => setState(() {}));
                      },
                    ),
                  ),
                ),
                Expanded(
                    flex: 8,
                    child: Center(
                        child: Text(
                      DateTextControllerFormatted,
                      style: Theme.of(context)
                          .textTheme
                          .titleMedium!
                          .copyWith(fontWeight: FontWeight.bold),
                    ))),
                Expanded(
                  flex: 1,
                  child: Align(
                    alignment: Alignment.centerRight,
                    child: IconButton(
                      icon: const Icon(Icons.calendar_month),
                      onPressed: () async {
                        final datePicked = await showDatePicker(
                                context: context,
                                firstDate: START_DATE,
                                lastDate: FINAL_DATE,
                                initialDate: DateTextController) ??
                            DateTime.now();

                        DateTextController = DateTime.utc(datePicked.year, datePicked.month, datePicked.day);
                        final TIME_DIFF = DateTextController.difference(START_DATE).inHours/24;
                        print("Difference: $TIME_DIFF");
                        old_page =
                            DateTextController.difference(START_DATE).inDays;
                        page_controller.jumpToPage(
                            DateTextController.difference(START_DATE).inDays);
                      },
                    ),
                  ),
                ),
              ],
            ),
          ),
          const Divider(),
          Expanded(
            child: PageView.builder(
                itemCount: FINAL_DATE.difference(START_DATE).inDays + 1,
                controller: page_controller,
                onPageChanged: (new_page) {
                  if (new_page < old_page) {
                    DateTextController = DateTime(DateTextController.year,
                        DateTextController.month, DateTextController.day - 1);
                  }
                  if (new_page > old_page) {
                    DateTextController = DateTime(DateTextController.year,
                        DateTextController.month, DateTextController.day + 1);
                  }
                  Format_Date();
                  old_page = new_page;
                  setState(() {});
                },
                itemBuilder: (context, position) {
                  return ListView(
                      padding: const EdgeInsets.symmetric(horizontal: 10),
                      children: FoodList(position, context));
                }),
          ),
        ],
      ),
    );
  }
}

// Calculates which day of the cycle the food should be based on the date
Find_Food_List(int day_pos) {
  int day = 0;
  int week = 0;
  Week week_list = week_1_2023T2;
  print("DAYS POS: $day_pos");
  if (day_pos <= MID_DATE.difference(START_DATE).inDays) {
    final days_diff = day_pos;
    final days_mod = days_diff % 21;
    week = ((((days_mod - (days_mod % 7)) / 7 + 1).toInt() + (FIRST_TERM_START_WEEK - 2)) % 3) + 1;
    day = (days_mod % 7) + 1;
    print("The term is 1 and the week is $week and the day is $day");
    switch (week) {
      case (1):
        week_list = FIRST_TERM.Week_1;
        // print("First Term Week 1");
        break;
      case (2):
        week_list = FIRST_TERM.Week_2;
        // print("First Term Week 2");
        break;
      case (3):
        week_list = FIRST_TERM.Week_3;
        // print("First Term Week 3");
        break;
      default:
    }
  } else {
    final days_diff = day_pos - MID_DATE.difference(START_DATE).inDays - 1;
    final days_mod = days_diff % 21;
    week = ((((days_mod - (days_mod % 7)) / 7 + 1).toInt() + (SECOND_TERM_START_WEEK - 2)) % 3) + 1;
    day = (days_mod % 7) + 1;
    print("The term is 2 and the week is $week and the day is $day");
    switch (week) {
      case (1):
        week_list = SECOND_TERM.Week_1;
        // print("Second Term Week 1");
        break;
      case (2):
        week_list = SECOND_TERM.Week_2;
        // print("Second Term Week 2");
        break;
      case (3):
        week_list = SECOND_TERM.Week_3;
        // print("Second Term Week 3");
        break;
      default:
    }
  }
  switch (day) {
    case (1):
      return week_list.Monday;
    case (2):
      return week_list.Tuesday;
    case (3):
      return week_list.Wednesday;
    case (4):
      return week_list.Thursday;
    case (5):
      return week_list.Friday;
    case (6):
      return week_list.Saturday;
    case (7):
      return week_list.Sunday;
    default:
  }
}

List<Widget> FoodList(int day_pos, BuildContext context) {
  ThemeData theme = Theme.of(context);
  List<FoodItem> food_list = Find_Food_List(day_pos);
  var widgetList = <Widget>[];

  widgetList.add(const SizedBox(height: 10));
  widgetList.add(Center(
      child: Text("Breakfast (7:30 am - 10 am)",
          style: theme.textTheme.titleMedium!
              .copyWith(fontWeight: FontWeight.w800, fontSize: 16.5))));
  widgetList.add(const SizedBox(height: 4.5));
  for (var breakfastFood in food_list) {
    if (breakfastFood.breakfast == true &&
        Include_Food(breakfastFood) == true) {
      // Add space between items
      widgetList.add(const SizedBox(height: 8.5));
      // Add list item
      widgetList.add(FoodListItem(breakfastFood));
    }
  }
  if ((day_pos % 7) == 5 || (day_pos % 7) == 6) {
    widgetList.add(const SizedBox(height: 13));
    widgetList.add(Center(
        child: Text("Brunch (10 am - 12 pm)",
            style: theme.textTheme.titleMedium!
                .copyWith(fontWeight: FontWeight.w800, fontSize: 16.5))));
    widgetList.add(const SizedBox(height: 4.5));
    for (var brunchFood in food_list) {
      if (brunchFood.brunch == true && Include_Food(brunchFood) == true) {
        // Add space between items
        widgetList.add(const SizedBox(height: 8.5));
        // Add list item
        widgetList.add(FoodListItem(brunchFood));
      }
    }
  }
  widgetList.add(const SizedBox(height: 13));
  widgetList.add(Center(
      child: Text("Lunch (12 pm - 2:15 pm)",
          style: theme.textTheme.titleMedium!
              .copyWith(fontWeight: FontWeight.w800, fontSize: 16.5))));
  widgetList.add(const SizedBox(height: 4.5));
  for (var lunchFood in food_list) {
    if (lunchFood.lunch == true && Include_Food(lunchFood) == true) {
      // Add space between items
      if (lunchFood.subtitle == false) {
        widgetList.add(const SizedBox(height: 8.5));
      }
      // Add list item
      widgetList.add(FoodListItem(lunchFood));
    }
  }
  widgetList.add(const SizedBox(height: 13));
  widgetList.add(Center(
      child: Text("Dinner (5 pm - 7:30 pm)",
          style: theme.textTheme.titleMedium!
              .copyWith(fontWeight: FontWeight.w800, fontSize: 16.5))));
  widgetList.add(const SizedBox(height: 4.5));
  for (var dinnerFood in food_list) {
    if (dinnerFood.dinner == true && Include_Food(dinnerFood) == true) {
      // Add space between items
      if (dinnerFood.subtitle == false) {
        widgetList.add(const SizedBox(height: 8.5));
      }
      // Add list item
      widgetList.add(FoodListItem(dinnerFood));
    }
  }
  widgetList.add(const SizedBox(height: 10));
  for (var dessertFood in food_list) {
    if (dessertFood.dessert == true && Include_Food(dessertFood) == true) {
      // Add space between items
      widgetList.add(const SizedBox(height: 14));
      // Add list item
      widgetList.add(FoodListItem(dessertFood));
    }
  }
  widgetList.add(const SizedBox(height: 10));
  return widgetList;
}

class FoodListItem extends StatelessWidget {
  const FoodListItem(this.Food);
  final FoodItem Food;

  @override
  Widget build(BuildContext context) {
    ThemeData theme = Theme.of(context);
    if (Food.title == true) {
      return Text(
        textAlign: TextAlign.center,
        style: theme.textTheme.bodyMedium!
            .copyWith(fontSize: 13.5, fontWeight: FontWeight.w600),
        Food.text,
      );
    } else if (Food.subtitle == true) {
      return Text(
        textAlign: TextAlign.center,
        style: theme.textTheme.bodyMedium!
            .copyWith(fontSize: 13.5, fontWeight: FontWeight.w400),
        Food.text,
      );
    } else {
      return Text(
        textAlign: TextAlign.center,
        style: theme.textTheme.bodyMedium!.copyWith(fontSize: 13),
        FoodCorrectText(Food),
      );
    }
  }
}

String FoodCorrectText(FoodItem Food) {
  if (setting_hide_dietary_tags == true) {
    return Food.text;
  }
  final food_text = Food.text;
  final List<bool> food_info_bool = [
    Food.dairy_free,
    Food.gluten_free,
    Food.low_gluten,
    Food.veg
  ];
  final List<String> food_info_string = ["DF", "GF", "LG", "VG"];
  bool comma_needed = false;
  String food_info = "";

  for (int i = 0; i < 4; i++) {
    if (food_info_bool[i] == true) {
      final food_info_item = food_info_string[i];
      if (comma_needed == true) {
        food_info = "$food_info, ";
      }
      food_info = "$food_info$food_info_item";
      comma_needed = true;
    }
  }

  if (comma_needed == true) {
    return "$food_text ($food_info)";
  } else {
    return food_text;
  }
}

bool Include_Food(FoodItem Food) {
  bool include = true;
  if (Food.title == true || Food.subtitle == true || Food.override == true) {
    return include;
  }
  if (setting_veg == true && Food.veg == false) {
    include = false;
  } else if (setting_dairy_allergy == true && Food.dairy_free == false) {
    include = false;
  } else if (setting_gluten_allergy == true &&
      Food.gluten_free == false &&
      Food.low_gluten == false) {
    include = false;
  }
  return include;
}

///////////////////////////////////////////////////////////////////////
////////////////////////// Settings Page //////////////////////////////
///////////////////////////////////////////////////////////////////////

class SettingsPage extends StatefulWidget {
  const SettingsPage({super.key, required this.ChangeDarkMode});
  final Function ChangeDarkMode;

  @override
  State<SettingsPage> createState() => _SettingsPageState();
}

class _SettingsPageState extends State<SettingsPage> {
  @override
  Widget build(BuildContext context) {
    return Material(
      child: Scaffold(
        appBar: AppBar(
          centerTitle: true,
          title: const Text("Settings"),
        ),
        body: Column(children: [
          Expanded(
              child: ScrollConfiguration(
                  behavior: MyBehavior(),
                  child: ListView(
                      padding: const EdgeInsets.symmetric(horizontal: 10),
                      children: [
                        const SizedBox(height: 15),
                        const Text("App Settings:"),
                        const SizedBox(height: 10),
                        InkWell(
                          onTap: () {
                            setState(() {
                              setting_dark_mode = !setting_dark_mode;
                              Update_settings();
                              widget.ChangeDarkMode();
                            });
                          },
                          child: Row(
                            children: [
                              IgnorePointer(
                                child: Checkbox(
                                  value: setting_dark_mode,
                                  onChanged: (bool? value) {
                                  },
                                ),
                              ),
                              const Center(
                                child: Text("Dark Mode"),
                              ),
                            ],
                          ),
                        ),
                        const SizedBox(height: 10),
                        InkWell(
                          onTap: () {
                            setState(() {
                              setting_hide_dietary_tags =
                                  !setting_hide_dietary_tags;
                              Update_settings();
                            });
                          },
                          child: Row(
                            children: [
                              IgnorePointer(
                                child: Checkbox(
                                  value: setting_hide_dietary_tags,
                                  onChanged: (bool? value) {},
                                ),
                              ),
                              const Center(
                                child: Text("Hide Dietary Tags"),
                              ),
                            ],
                          ),
                        ),
                        const Divider(),
                        const SizedBox(height: 10),
                        const Text("Dietary Requirements:"),
                        const SizedBox(height: 10),
                        InkWell(
                          onTap: () {
                            setState(() {
                              setting_veg = !setting_veg;
                              Update_settings();
                            });
                          },
                          child: Row(
                            children: [
                              IgnorePointer(
                                child: Checkbox(
                                  value: setting_veg,
                                  onChanged: (bool? value) {},
                                ),
                              ),
                              const Expanded(
                                child: Text(
                                    "Vegetarian (excludes non vegetarian foods)"),
                              ),
                            ],
                          ),
                        ),
                        const SizedBox(height: 10),
                        InkWell(
                          onTap: () {
                            setState(() {
                              setting_dairy_allergy = !setting_dairy_allergy;
                              Update_settings();
                            });
                          },
                          child: Row(
                            children: [
                              IgnorePointer(
                                child: Checkbox(
                                  value: setting_dairy_allergy,
                                  onChanged: (bool? value) {},
                                ),
                              ),
                              const Expanded(
                                child: Text(
                                    "Dairy Allergy (BETA - excludes non DF foods)"),
                              ),
                            ],
                          ),
                        ),
                        const SizedBox(height: 10),
                        InkWell(
                          onTap: () {
                            setState(() {
                              setting_gluten_allergy = !setting_gluten_allergy;
                              Update_settings();
                            });
                          },
                          child: Row(
                            children: [
                              IgnorePointer(
                                child: Checkbox(
                                  value: setting_gluten_allergy,
                                  onChanged: (bool? value) {},
                                ),
                              ),
                              Expanded(
                                child: Text(
                                    "Gluten Allergy (BETA - excludes non GF & LG foods)"),
                              ),
                            ],
                          ),
                        ),
                        const Divider(),
                        const SizedBox(height: 10),
                        RichText(
                          text: TextSpan(
                            style: Theme.of(context)
                                .textTheme
                                .titleMedium!
                                .copyWith(
                                    fontSize: 15, fontWeight: FontWeight.w500),
                            children: [
                              const TextSpan(
                                  text:
                                      'If you find any bugs in the app or any major problem with the menu, feel free to reach out to me on '),
                              TextSpan(
                                text: 'messenger',
                                style: const TextStyle(
                                    color: Colors.lightBlue,
                                    fontWeight: FontWeight.bold),
                                recognizer: TapGestureRecognizer()
                                  ..onTap = () async {
                                    final url = Uri.https(
                                        'www.facebook.com',
                                        '/profile.php',
                                        {'id': '100090069107079'});
                                    if (await canLaunchUrl(url)) {
                                      await launchUrl(url);
                                    }
                                  },
                              ),
                              const TextSpan(text: '.'),
                            ],
                          ),
                        ),
                        const SizedBox(height: 10)
                      ]))),
          Container(
              margin: EdgeInsets.all(8.0),
              width: double
                  .infinity, // This makes the container take up all available horizontal space
              child: ElevatedButton(
                style: ElevatedButton.styleFrom(
                    backgroundColor: setting_dark_mode == true
                        ? Color.lerp(
                            Theme.of(context).primaryColor, Colors.white, 0.2)
                        : Theme.of(context).primaryColor,
                    foregroundColor: Theme.of(context).colorScheme.onPrimary
                    ),
                onPressed: () async {
                  final url = Uri.parse(
                      "https://login.microsoftonline.com/common/oauth2/authorize?client_id=c9a559d2-7aab-4f13-a6ed-e7e9c52aec87&resource=c9a559d2-7aab-4f13-a6ed-e7e9c52aec87&response_type=code%20id_token&scope=openid%20profile&state=OpenIdConnect.AuthenticationProperties%3DeyJ2ZXJzaW9uIjoxLCJkYXRhIjp7IklkZW50aXR5UHJvdmlkZXIiOiJBYkpVT0hSSEMyTEhLS25LeTk4bjVCazBDYlh4UVJaclhVQWQtSFBKUENCaloxZnBKZ0Fxa0FCTGo0Zjc3NnFIZGhTTDNob0tqX2tlQW42bFFpNEhfck0iLCIucmVkaXJlY3QiOiJodHRwczovL2Zvcm1zLm9mZmljZS5jb20vUGFnZXMvUmVzcG9uc2VQYWdlLmFzcHg_aWQ9cE1fMlB4WG4yMGk0NFFobnVmbjdvOTFEWVVRNmxXOU1zR0xrOGFWOUFnTlVObEZYVERVd1VFZ3dWekpRTlVWWVJqZE1RVmRKTmt4U01TNHUmb3JpZ2luPVFSQ29kZSZzaWQ9ZjdkMjc1ZTQtMTg4MS00YTMyLWI5NTgtYzNjZTc5YTI5Mzc4In19&response_mode=form_post&nonce=638204498873498389.OWY4MTQ1YzYtMWJiMi00NGY5LTg5MjUtODI2Y2E2OTE0NTVkMzBlMDYxMjAtNDgxOC00MWVhLWI0ZWUtODc1YTg4MDBlMTc4&redirect_uri=https%3A%2F%2Fforms.office.com%2Flanding&msafed=0&x-client-SKU=ID_NET472&x-client-ver=6.16.0.0&sso_reload=true");
                  if (await canLaunchUrl(url)) {
                    await launchUrl(url);
                  }
                },
                child: Text('Feedback Form for Food',
                    style:
                        TextStyle(fontSize: 15, fontWeight: FontWeight.bold)),
              )),
        ]),
      ),
    );
  }
}

// https://login.microsoftonline.com/common/oauth2/authorize?client_id=c9a559d2-7aab-4f13-a6ed-e7e9c52aec87&resource=c9a559d2-7aab-4f13-a6ed-e7e9c52aec87&response_type=code%20id_token&scope=openid%20profile&state=OpenIdConnect.AuthenticationProperties%3DeyJ2ZXJzaW9uIjoxLCJkYXRhIjp7IklkZW50aXR5UHJvdmlkZXIiOiJBYkpVT0hSSEMyTEhLS25LeTk4bjVCazBDYlh4UVJaclhVQWQtSFBKUENCaloxZnBKZ0Fxa0FCTGo0Zjc3NnFIZGhTTDNob0tqX2tlQW42bFFpNEhfck0iLCIucmVkaXJlY3QiOiJodHRwczovL2Zvcm1zLm9mZmljZS5jb20vUGFnZXMvUmVzcG9uc2VQYWdlLmFzcHg_aWQ9cE1fMlB4WG4yMGk0NFFobnVmbjdvOTFEWVVRNmxXOU1zR0xrOGFWOUFnTlVObEZYVERVd1VFZ3dWekpRTlVWWVJqZE1RVmRKTmt4U01TNHUmb3JpZ2luPVFSQ29kZSZzaWQ9ZjdkMjc1ZTQtMTg4MS00YTMyLWI5NTgtYzNjZTc5YTI5Mzc4In19&response_mode=form_post&nonce=638204498873498389.OWY4MTQ1YzYtMWJiMi00NGY5LTg5MjUtODI2Y2E2OTE0NTVkMzBlMDYxMjAtNDgxOC00MWVhLWI0ZWUtODc1YTg4MDBlMTc4&redirect_uri=https%3A%2F%2Fforms.office.com%2Flanding&msafed=0&x-client-SKU=ID_NET472&x-client-ver=6.16.0.0&sso_reload=true

class MyBehavior extends ScrollBehavior {
  @override
  Widget buildOverscrollIndicator(
      BuildContext context, Widget child, ScrollableDetails details) {
    return child;
  }
}
