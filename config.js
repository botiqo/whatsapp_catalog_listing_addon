// Constants
const HEADERS = ["id", "thumbnail", "price", "name", "description", "image_url", "retailer_id", "brand", "variant_group_id", "url", "currency", "category_id", "availability", "condition", "sale_price", "is_hidden"];
const CUSTOM_COLUMNS = ["thumbnail"];
const PRODUCT_TYPE_COLUMNS = {
  'standard': ['id', 'thumbnail', 'description', 'name', 'price', 'currency', 'image_url', 'availability', 'condition', 'brand', 'category_id', 'url', 'retailer_id'],
  'service': ['id', 'thumbnail', 'description', 'name', 'price', 'currency', 'image_url', 'availability', 'category_id', 'url'],
  'default': HEADERS
};
const CURRENCY_LIST = ["ILS", "AED", "USD", "CAD", "EUR", "GBP", "INR", "MXN", "BRL", "IDR", "ZAR"];
const CATEGORY_LIST = [
  "AUTO_VEHICLES_PARTS_ACCESSORIES",
  "BEAUTY_HEALTH_HAIR",
  "BUSINESS_SERVICES",
  "BABY_KIDS_GOODS",
  "COMMERCIAL_EQUIPMENT",
  "ELECTRONICS",
  "FOOD_BEVERAGES",
  "FURNITURE_APPLIANCES",
  "HOME_GOODS_DECOR",
  "LUGGAGE_BAGS",
  "MEDIA_MUSIC_BOOKS",
  "MISC",
  "PERSONAL_ACCESSORIES",
  "PET_SUPPLIES",
  "SPORTING_GOODS",
  "TOYS_GAMES_COLLECTIBLES",
  "APPAREL_ACCESSORIES",
  "FOOTWEAR",
  "HAIR_EXTENSIONS_WIGS",
  "HAIR_STYLING_TOOLS",
  "MAKEUP_COSMETICS",
  "FRAGRANCES",
  "SKIN_CARE",
  "BATH_BODY",
  "NAIL_CARE",
  "VITAMINS_SUPPLEMENTS",
  "MEDICAL_SUPPLIES_EQUIPMENT",
  "TICKETS",
  "TRAVEL_SERVICES"
];
const PRODUCT_TYPE_LIST = ["standard", "service", "default"];
const AVAILABILITY_LIST = ["in stock", "out of stock"];
const CONDITION_LIST = ["new", "used"];
const LOADER_IMG_URL = "https://example.com/loading.gif"; // Replace with an actual loading GIF
/**
 * Retrieves configuration dropdown lists and preselected values.
 * @return {Object} An object containing dropdown lists and preselected values.
 */
function getConfigurationDropdownLists() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const configData = {
      currencyList: CURRENCY_LIST,
      categoryList: CATEGORY_LIST,
      productTypeList: PRODUCT_TYPE_LIST,
      availabilityList: AVAILABILITY_LIST,
      conditionList: CONDITION_LIST,
      preselectedProductType: scriptProperties.getProperty('DEFAULT_PRODUCT_TYPE') || "",
      preselectedCurrency: scriptProperties.getProperty('DEFAULT_CURRENCY') || "",
      preselectedCategory: scriptProperties.getProperty('DEFAULT_CATEGORY') || "",
      preselectedAvailability: scriptProperties.getProperty('DEFAULT_AVAILABILITY') || "",
      preselectedCondition: scriptProperties.getProperty('DEFAULT_CONDITION') || ""
    };

    ErrorHandler.log('Configuration dropdown lists retrieved', 'INFO');
    return configData;
  } catch (error) {
    ErrorHandler.handleError(error, "Error Please try again or contact support.");
    throw error;
  }
}