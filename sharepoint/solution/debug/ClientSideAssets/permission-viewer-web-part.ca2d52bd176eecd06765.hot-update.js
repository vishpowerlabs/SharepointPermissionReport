"use strict";
self["webpackHotUpdate_76d5422e_a579_4a00_99b3_192667104868_0_0_1"]("permission-viewer-web-part",{

/***/ 89135:
/*!**********************************************************************!*\
  !*** ./lib/webparts/permissionViewer/components/PermissionViewer.js ***!
  \**********************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   "default": () => (__WEBPACK_DEFAULT_EXPORT__)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_9__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! react */ 85959);
/* harmony import */ var react__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(react__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_11__ = __webpack_require__(/*! @fluentui/react/lib/Pivot */ 92070);
/* harmony import */ var _fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__ = __webpack_require__(/*! @fluentui/react/lib/Pivot */ 15369);
/* harmony import */ var _fluentui_react_lib_SearchBox__WEBPACK_IMPORTED_MODULE_13__ = __webpack_require__(/*! @fluentui/react/lib/SearchBox */ 21262);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 5613);
/* harmony import */ var _fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__ = __webpack_require__(/*! @fluentui/react/lib/Button */ 29425);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_17__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 80954);
/* harmony import */ var _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_18__ = __webpack_require__(/*! @fluentui/react/lib/Spinner */ 49885);
/* harmony import */ var _services_PermissionService__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ../services/PermissionService */ 94110);
/* harmony import */ var _Header__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! ./Header */ 26783);
/* harmony import */ var _StatsCards__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! ./StatsCards */ 54088);
/* harmony import */ var _SitePermissions__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! ./SitePermissions */ 82561);
/* harmony import */ var _ListPermissions__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! ./ListPermissions */ 20548);
/* harmony import */ var _LoadingState__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! ./LoadingState */ 37087);
/* harmony import */ var _DeepScanDialog__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./DeepScanDialog */ 34937);
/* harmony import */ var _utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__ = __webpack_require__(/*! ../utils/CsvExport */ 2759);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_15__ = __webpack_require__(/*! @fluentui/react */ 4312);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_16__ = __webpack_require__(/*! @fluentui/react */ 10548);
/* harmony import */ var _fluentui_react__WEBPACK_IMPORTED_MODULE_19__ = __webpack_require__(/*! @fluentui/react */ 87295);
/* harmony import */ var _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./PermissionViewer.module.scss */ 29929);
















var PermissionViewer = function (props) {
    var _a;
    var _b = react__WEBPACK_IMPORTED_MODULE_0__.useState('site'), activeTab = _b[0], setActiveTab = _b[1];
    var _c = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), sitePermissions = _c[0], setSitePermissions = _c[1];
    var _d = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredSitePermissions = _d[0], setFilteredSitePermissions = _d[1];
    var _e = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), lists = _e[0], setLists = _e[1];
    var _f = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), filteredLists = _f[0], setFilteredLists = _f[1];
    var _g = react__WEBPACK_IMPORTED_MODULE_0__.useState({ totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 }), stats = _g[0], setStats = _g[1];
    var _h = react__WEBPACK_IMPORTED_MODULE_0__.useState(true), isLoading = _h[0], setIsLoading = _h[1];
    var _j = react__WEBPACK_IMPORTED_MODULE_0__.useState('Loading site permissions...'), loadingMessage = _j[0], setLoadingMessage = _j[1];
    var _k = react__WEBPACK_IMPORTED_MODULE_0__.useState(), permissionService = _k[0], setPermissionService = _k[1];
    var _l = react__WEBPACK_IMPORTED_MODULE_0__.useState(''), searchText = _l[0], setSearchText = _l[1];
    var _m = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isExporting = _m[0], setIsExporting = _m[1];
    var _o = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isScanning = _o[0], setIsScanning = _o[1];
    // Deep Scan State
    var _p = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), isDeepScanOpen = _p[0], setIsDeepScanOpen = _p[1];
    var _q = react__WEBPACK_IMPORTED_MODULE_0__.useState([]), deepScanItems = _q[0], setDeepScanItems = _q[1];
    var _r = react__WEBPACK_IMPORTED_MODULE_0__.useState(''), deepScanListTitle = _r[0], setDeepScanListTitle = _r[1];
    var _s = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), confirmScanList = _s[0], setConfirmScanList = _s[1];
    var _t = react__WEBPACK_IMPORTED_MODULE_0__.useState(false), scanNoResults = _t[0], setScanNoResults = _t[1];
    var _u = react__WEBPACK_IMPORTED_MODULE_0__.useState(null), errorMessage = _u[0], setErrorMessage = _u[1];
    // Delete Confirmation State
    var _v = react__WEBPACK_IMPORTED_MODULE_0__.useState({ isOpen: false, title: '', subText: '', onConfirm: function () { } }), deleteConfirmState = _v[0], setDeleteConfirmState = _v[1];
    react__WEBPACK_IMPORTED_MODULE_0__.useEffect(function () {
        var service = new _services_PermissionService__WEBPACK_IMPORTED_MODULE_1__.PermissionServiceImpl(props.spHttpClient, props.webUrl);
        setPermissionService(service);
        loadData(service);
    }, [props.excludedLists]);
    var loadData = function (service) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var sitePerms, siteStats, listsData, uniqueListsCount;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setIsLoading(true);
                    setLoadingMessage('Loading site permissions...');
                    return [4 /*yield*/, service.getSiteRoleAssignments()];
                case 1:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    setFilteredSitePermissions(sitePerms);
                    return [4 /*yield*/, service.getSiteStats()];
                case 2:
                    siteStats = _a.sent();
                    // Load lists
                    setLoadingMessage('Loading lists and libraries...');
                    return [4 /*yield*/, service.getLists(props.excludedLists)];
                case 3:
                    listsData = _a.sent();
                    uniqueListsCount = listsData.filter(function (l) { return l.HasUniqueRoleAssignments; }).length;
                    setStats((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, siteStats), { uniquePermissionsCount: uniqueListsCount + (sitePerms.length > 0 ? 1 : 0) // Simplified logic: +1 if site has perms (it always does)
                     }));
                    // Sort lists: Unique first, then Inherited
                    listsData.sort(function (a, b) {
                        if (a.HasUniqueRoleAssignments === b.HasUniqueRoleAssignments)
                            return 0;
                        return a.HasUniqueRoleAssignments ? -1 : 1;
                    });
                    setLists(listsData);
                    setFilteredLists(listsData);
                    setIsLoading(false);
                    return [2 /*return*/];
            }
        });
    }); };
    var handleRefresh = function () {
        if (permissionService) {
            loadData(permissionService);
        }
    };
    var onSearch = function (newValue) {
        setSearchText(newValue);
        var lower = newValue.toLowerCase();
        if (activeTab === 'site') {
            var filtered = sitePermissions.filter(function (p) {
                var _a;
                return p.Member.Title.toLowerCase().includes(lower) ||
                    ((_a = p.Member.Email) === null || _a === void 0 ? void 0 : _a.toLowerCase().includes(lower));
            });
            setFilteredSitePermissions(filtered);
        }
        else {
            var filtered = lists.filter(function (l) {
                return l.Title.toLowerCase().includes(lower) ||
                    l.ServerRelativeUrl.toLowerCase().includes(lower);
            });
            setFilteredLists(filtered);
        }
    };
    var handleGetListPermissions = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/, []];
                    list = lists.find(function (l) { return l.Id === listId; });
                    return [4 /*yield*/, permissionService.getListRoleAssignments(listId, list ? list.Title : '')];
                case 1: return [2 /*return*/, _a.sent()];
            }
        });
    }); };
    var handleRemoveSitePermission = function (principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', "? This will remove all permissions for this user on this site."),
            onConfirm: function () { return executeRemoveSitePermission(principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveSitePermission = function (principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var success, sitePerms, lower_1, filtered, error_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                    setIsLoading(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    return [4 /*yield*/, permissionService.removeSitePermission(principalId)];
                case 2:
                    success = _a.sent();
                    if (!success) return [3 /*break*/, 4];
                    return [4 /*yield*/, permissionService.getSiteRoleAssignments()];
                case 3:
                    sitePerms = _a.sent();
                    setSitePermissions(sitePerms);
                    if (searchText) {
                        lower_1 = searchText.toLowerCase();
                        filtered = sitePerms.filter(function (p) {
                            var _a;
                            return p.Member.Title.toLowerCase().includes(lower_1) ||
                                ((_a = p.Member.Email) === null || _a === void 0 ? void 0 : _a.toLowerCase().includes(lower_1));
                        });
                        setFilteredSitePermissions(filtered);
                    }
                    else {
                        setFilteredSitePermissions(sitePerms);
                    }
                    return [3 /*break*/, 5];
                case 4:
                    setErrorMessage("Failed to remove permission. Please try again or check console for details.");
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_1 = _a.sent();
                    console.error("Error removing site permission", error_1);
                    setErrorMessage("An unexpected error occurred while removing permission.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsLoading(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    var handleRemoveListPermission = function (listId, principalId, principalName) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            return [2 /*return*/, new Promise(function (resolve) {
                    setDeleteConfirmState({
                        isOpen: true,
                        title: "Remove Permissions?",
                        subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this list?"),
                        onConfirm: function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
                            var success, error_2;
                            return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                                        if (!permissionService) {
                                            resolve(false);
                                            return [2 /*return*/];
                                        }
                                        _a.label = 1;
                                    case 1:
                                        _a.trys.push([1, 3, , 4]);
                                        return [4 /*yield*/, permissionService.removeListPermission(listId, principalId)];
                                    case 2:
                                        success = _a.sent();
                                        if (!success) {
                                            setErrorMessage("Failed to remove list permission. Please try again or check console for details.");
                                        }
                                        resolve(success);
                                        return [3 /*break*/, 4];
                                    case 3:
                                        error_2 = _a.sent();
                                        console.error("Error removing list permission", error_2);
                                        setErrorMessage("An unexpected error occurred while removing list permission.");
                                        resolve(false);
                                        return [3 /*break*/, 4];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); },
                        onCancel: function () {
                            resolve(false);
                        }
                    });
                })];
        });
    }); };
    var handleRemoveDeepScanItemPermission = function (itemId, principalId, principalName) {
        setDeleteConfirmState({
            isOpen: true,
            title: "Remove Permissions?",
            subText: "Are you sure you want to remove permissions for ".concat(principalName || 'this user', " on this item?"),
            onConfirm: function () { return executeRemoveDeepScanItemPermission(itemId, principalId); },
            onCancel: function () { }
        });
    };
    var executeRemoveDeepScanItemPermission = function (itemId, principalId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list, success, error_3;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !deepScanListTitle)
                        return [2 /*return*/];
                    setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); });
                    list = lists.find(function (l) { return l.Title === deepScanListTitle; });
                    if (!list)
                        return [2 /*return*/];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, permissionService.removeItemPermission(list.Id, itemId, principalId)];
                case 2:
                    success = _a.sent();
                    if (success) {
                        setDeepScanItems(function (prevItems) {
                            return prevItems.map(function (item) {
                                if (item.Id === itemId) {
                                    var newRoles = item.RoleAssignments.filter(function (ra) { return ra.PrincipalId !== principalId; });
                                    if (newRoles.length === 0) {
                                        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, item), { RoleAssignments: [] });
                                    }
                                    return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, item), { RoleAssignments: newRoles });
                                }
                                return item;
                            }).filter(function (item) { return item !== null; });
                        });
                    }
                    else {
                        setErrorMessage("Failed to remove item permission.");
                    }
                    return [3 /*break*/, 4];
                case 3:
                    error_3 = _a.sent();
                    console.error("Error removing item permission", error_3);
                    setErrorMessage("An unexpected error occurred while removing item permission.");
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var handleExport = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var error_4;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService)
                        return [2 /*return*/];
                    setIsExporting(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 6, 7, 8]);
                    if (!(activeTab === 'site')) return [3 /*break*/, 3];
                    return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportSitePermissions)(filteredSitePermissions, permissionService)];
                case 2:
                    _a.sent();
                    return [3 /*break*/, 5];
                case 3: return [4 /*yield*/, (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportListPermissions)(filteredLists, permissionService)];
                case 4:
                    _a.sent();
                    _a.label = 5;
                case 5: return [3 /*break*/, 8];
                case 6:
                    error_4 = _a.sent();
                    console.error("Export failed", error_4);
                    alert("Export failed. See console for details.");
                    return [3 /*break*/, 8];
                case 7:
                    setIsExporting(false);
                    return [7 /*endfinally*/];
                case 8: return [2 /*return*/];
            }
        });
    }); };
    // ...
    var handleDeepScan = function (listId) { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var list;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            list = lists.find(function (l) { return l.Id === listId; });
            if (!list)
                return [2 /*return*/];
            setScanNoResults(false);
            setConfirmScanList({ id: list.Id, title: list.Title });
            return [2 /*return*/];
        });
    }); };
    var executeDeepScan = function () { return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__awaiter)(void 0, void 0, void 0, function () {
        var listId, listTitle, items, e_1;
        return (0,tslib__WEBPACK_IMPORTED_MODULE_9__.__generator)(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!permissionService || !confirmScanList)
                        return [2 /*return*/];
                    listId = confirmScanList.id;
                    listTitle = confirmScanList.title;
                    setIsScanning(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, permissionService.getUniquePermissionItems(listId)];
                case 2:
                    items = _a.sent();
                    if (items.length > 0) {
                        setDeepScanItems(items);
                        setDeepScanListTitle(listTitle);
                        setIsDeepScanOpen(true);
                        setConfirmScanList(null); // Close confirm dialog only if opening results
                    }
                    else {
                        setScanNoResults(true);
                        // Keep dialog open to show message
                    }
                    return [3 /*break*/, 5];
                case 3:
                    e_1 = _a.sent();
                    console.error(e_1);
                    alert("Error during deep scan.");
                    setConfirmScanList(null); // Close on error too
                    return [3 /*break*/, 5];
                case 4:
                    setIsScanning(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var downloadDeepScanResults = function () {
        (0,_utils_CsvExport__WEBPACK_IMPORTED_MODULE_10__.exportDeepScanResults)(deepScanItems, deepScanListTitle);
    };
    var getDialogTitle = function () {
        if (scanNoResults)
            return 'Deep Scan Complete';
        if (isScanning)
            return 'Deep Scan in Progress...';
        return 'Start Deep Scan?';
    };
    var getDialogSubText = function () {
        if (scanNoResults) {
            return "No items with unique permissions were found in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". All items inherit permissions.");
        }
        if (isScanning) {
            return "Scanning \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\". This may take a few moments depending on the number of items...");
        }
        return "This will verify every single item in \"".concat(confirmScanList === null || confirmScanList === void 0 ? void 0 : confirmScanList.title, "\" to find unique permissions. This might take a while for large lists. Continue?");
    };
    return (react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].permissionViewer },
        react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].webpartContainer },
            (props.showComponentHeader !== false) && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_Header__WEBPACK_IMPORTED_MODULE_2__.Header, { onRefresh: handleRefresh, isLoading: isLoading, themeVariant: props.themeVariant, opacity: (_a = props.headerOpacity) !== null && _a !== void 0 ? _a : 100, title: props.webPartTitle, titleFontSize: props.webPartTitleFontSize })),
            (props.showStats !== false) && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_StatsCards__WEBPACK_IMPORTED_MODULE_3__.StatsCards, { stats: stats }),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].tabsContainer },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_11__.Pivot, { onLinkClick: function (item) {
                        if (item === null || item === void 0 ? void 0 : item.props.itemKey) {
                            setActiveTab(item.props.itemKey);
                        }
                    }, selectedKey: activeTab },
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__.PivotItem, { headerText: "Site Permissions", itemKey: "site", itemIcon: "Shield" }),
                    react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Pivot__WEBPACK_IMPORTED_MODULE_12__.PivotItem, { headerText: "Lists & Libraries", itemKey: "lists", itemIcon: "List" }))),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].toolbar },
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_SearchBox__WEBPACK_IMPORTED_MODULE_13__.SearchBox, { placeholder: "Search users or groups...", onChange: function (_, newValue) { return onSearch(newValue || ''); }, value: searchText, className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].searchBox }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { text: isExporting ? "Exporting..." : "Export to CSV", iconProps: { iconName: 'Download' }, onClick: handleExport, disabled: isExporting || isLoading, className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].exportBtn, styles: {
                        root: { height: '32px' }, // Maintain height
                        label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                    } })),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement("div", { className: _PermissionViewer_module_scss__WEBPACK_IMPORTED_MODULE_8__["default"].content },
                isLoading && react__WEBPACK_IMPORTED_MODULE_0__.createElement(_LoadingState__WEBPACK_IMPORTED_MODULE_6__.LoadingState, { message: loadingMessage }),
                !isLoading && activeTab === 'site' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_SitePermissions__WEBPACK_IMPORTED_MODULE_4__.SitePermissions, { permissions: filteredSitePermissions, permissionService: permissionService, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveSitePermission })),
                !isLoading && activeTab === 'lists' && (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_ListPermissions__WEBPACK_IMPORTED_MODULE_5__.ListPermissions, { lists: filteredLists, getListPermissions: handleGetListPermissions, onScanItems: handleDeepScan, themeVariant: props.themeVariant, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveListPermission })))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_DeepScanDialog__WEBPACK_IMPORTED_MODULE_7__.DeepScanDialog, { isOpen: isDeepScanOpen, onDismiss: function () { return setIsDeepScanOpen(false); }, listTitle: deepScanListTitle, items: deepScanItems, onDownload: downloadDeepScanResults, buttonFontSize: props.buttonFontSize, contentFontSize: props.contentFontSize, onRemovePermission: handleRemoveDeepScanItemPermission }),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !confirmScanList, onDismiss: function () { if (!isScanning)
                setConfirmScanList(null); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: getDialogTitle(),
                subText: getDialogSubText()
            } }, isScanning ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_17__.Spinner, { size: _fluentui_react_lib_Spinner__WEBPACK_IMPORTED_MODULE_18__.SpinnerSize.large, label: "Scanning for unique permissions..." })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null, scanNoResults ? (react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: function () { return setConfirmScanList(null); }, text: "Close", styles: {
                root: { height: '32px' },
                label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
            } })) : (react__WEBPACK_IMPORTED_MODULE_0__.createElement(react__WEBPACK_IMPORTED_MODULE_0__.Fragment, null,
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: executeDeepScan, text: "Start Scan", styles: {
                    root: { height: '32px' },
                    label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                } }),
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { onClick: function () { return setConfirmScanList(null); }, text: "Cancel", styles: {
                    root: { height: '32px' },
                    label: { fontSize: props.buttonFontSize || '12px', fontWeight: 600 }
                } })))))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !deleteConfirmState.isOpen, onDismiss: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); }); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: deleteConfirmState.title,
                subText: deleteConfirmState.subText,
            }, modalProps: {
                isBlocking: true,
                styles: { main: { maxWidth: 450 } }
            } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: deleteConfirmState.onConfirm, text: "Remove", styles: {
                        root: { background: '#d13438', border: '1px solid #d13438' }, // Red color for danger
                        rootHovered: { background: '#a4262c' },
                        label: { fontWeight: 600 }
                    } }),
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_14__.DefaultButton, { onClick: function () { return setDeleteConfirmState(function (prev) { return ((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_9__.__assign)({}, prev), { isOpen: false })); }); }, text: "Cancel" }))),
        react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_15__.Dialog, { hidden: !errorMessage, onDismiss: function () { return setErrorMessage(null); }, dialogContentProps: {
                type: _fluentui_react__WEBPACK_IMPORTED_MODULE_16__.DialogType.normal,
                title: 'Error',
                subText: errorMessage || 'An unexpected error occurred.'
            } },
            react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react__WEBPACK_IMPORTED_MODULE_19__.DialogFooter, null,
                react__WEBPACK_IMPORTED_MODULE_0__.createElement(_fluentui_react_lib_Button__WEBPACK_IMPORTED_MODULE_20__.PrimaryButton, { onClick: function () { return setErrorMessage(null); }, text: "OK" })))));
};
/* harmony default export */ const __WEBPACK_DEFAULT_EXPORT__ = (PermissionViewer);


/***/ }),

/***/ 94110:
/*!*********************************************************************!*\
  !*** ./lib/webparts/permissionViewer/services/PermissionService.js ***!
  \*********************************************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   PermissionServiceImpl: () => (/* binding */ PermissionServiceImpl)
/* harmony export */ });
/* harmony import */ var tslib__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! tslib */ 10196);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! @microsoft/sp-http */ 91909);
/* harmony import */ var _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(_microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__);


var PermissionServiceImpl = /** @class */ (function () {
    function PermissionServiceImpl(spHttpClient, webUrl) {
        this._spHttpClient = spHttpClient;
        this._webUrl = webUrl;
    }
    PermissionServiceImpl.prototype.getSiteRoleAssignments = function () {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint, response, json, error_1;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = "".concat(this._webUrl, "/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100");
                        return [4 /*yield*/, this._spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (json === null || json === void 0 ? void 0 : json.value) {
                            return [2 /*return*/, this._processRoleAssignments(json.value)];
                        }
                        return [2 /*return*/, []];
                    case 3:
                        error_1 = _a.sent();
                        console.error("Error fetching site permissions", error_1);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.getLists = function (excludedLists) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint, response, json, systemLists_1, error_2;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = "".concat(this._webUrl, "/_api/web/lists?$select=Id,Title,ItemCount,Hidden,BaseType,RootFolder/ServerRelativeUrl,HasUniqueRoleAssignments,EntityTypeName,BaseTemplate&$filter=Hidden eq false&$expand=RootFolder");
                        return [4 /*yield*/, this._spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (json === null || json === void 0 ? void 0 : json.value) {
                            systemLists_1 = new Set(excludedLists || [
                                "Site Assets", "Style Library", "Master Page Gallery",
                                "Form Templates", "User Information List", "Composed Looks", "Solution Gallery",
                                "TaxonomyHiddenList", "Appdata", "Appfiles"
                            ]);
                            return [2 /*return*/, json.value.filter(function (list) {
                                    // Exclude catalogs (BaseTemplate 113, 116 etc usually, but we check names/types)
                                    // BaseType 0 = Generic List, 1 = Document Library
                                    if (systemLists_1.has(list.Title))
                                        return false;
                                    if (list.EntityTypeName === "OData__catalogs")
                                        return false;
                                    return true;
                                }).map(function (list) {
                                    return {
                                        Id: list.Id,
                                        Title: list.Title,
                                        ItemCount: list.ItemCount,
                                        Hidden: list.Hidden,
                                        ItemType: list.BaseType === 1 ? 'Library' : 'List',
                                        ServerRelativeUrl: list.RootFolder ? list.RootFolder.ServerRelativeUrl : '',
                                        HasUniqueRoleAssignments: list.HasUniqueRoleAssignments,
                                        EntityTypeName: list.EntityTypeName
                                    };
                                })];
                        }
                        return [2 /*return*/, []];
                    case 3:
                        error_2 = _a.sent();
                        console.error("Error fetching lists", error_2);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.getListRoleAssignments = function (listId, listTitle) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint, response, json, error_3;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = "".concat(this._webUrl, "/_api/web/lists(guid'").concat(listId, "')/roleassignments?$expand=Member,RoleDefinitionBindings&$top=100");
                        return [4 /*yield*/, this._spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (json === null || json === void 0 ? void 0 : json.value) {
                            return [2 /*return*/, this._processRoleAssignments(json.value)];
                        }
                        return [2 /*return*/, []];
                    case 3:
                        error_3 = _a.sent();
                        console.error("Error fetching permissions for list ".concat(listTitle), error_3);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.getSiteStats = function () {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var usersEndpoint, groupsEndpoint, usersReq, groupsReq, _a, usersResp, groupsResp, usersJson, groupsJson, error_4;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 4, , 5]);
                        usersEndpoint = "".concat(this._webUrl, "/_api/web/siteusers");
                        groupsEndpoint = "".concat(this._webUrl, "/_api/web/sitegroups");
                        usersReq = this._spHttpClient.get(usersEndpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1);
                        groupsReq = this._spHttpClient.get(groupsEndpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1);
                        return [4 /*yield*/, Promise.all([usersReq, groupsReq])];
                    case 1:
                        _a = _b.sent(), usersResp = _a[0], groupsResp = _a[1];
                        return [4 /*yield*/, usersResp.json()];
                    case 2:
                        usersJson = _b.sent();
                        return [4 /*yield*/, groupsResp.json()];
                    case 3:
                        groupsJson = _b.sent();
                        return [2 /*return*/, {
                                totalUsers: usersJson.value ? usersJson.value.length : 0,
                                totalGroups: groupsJson.value ? groupsJson.value.length : 0,
                                uniquePermissionsCount: 0 // To be calculated by caller or separate logic
                            }];
                    case 4:
                        error_4 = _b.sent();
                        console.error("Error fetching stats", error_4);
                        return [2 /*return*/, { totalUsers: 0, totalGroups: 0, uniquePermissionsCount: 0 }];
                    case 5: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.getGroupMembers = function (groupId) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint, response, json, error_5;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 3, , 4]);
                        endpoint = "".concat(this._webUrl, "/_api/web/sitegroups/getbyid(").concat(groupId, ")/users");
                        return [4 /*yield*/, this._spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (json === null || json === void 0 ? void 0 : json.value) {
                            return [2 /*return*/, json.value];
                        }
                        return [2 /*return*/, []];
                    case 3:
                        error_5 = _a.sent();
                        console.error("Error fetching members for group ".concat(groupId), error_5);
                        return [2 /*return*/, []];
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.getUniquePermissionItems = function (listId) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint, response, json, allItems, uniqueItems, results_1, chunks, _i, chunks_1, chunk, promises, chunkResults, error_6;
            var _this = this;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 8, , 9]);
                        endpoint = "".concat(this._webUrl, "/_api/web/lists(guid'").concat(listId, "')/items?$select=Id,Title,FileRef,FileLeafRef,FileSystemObjectType,HasUniqueRoleAssignments&$top=5000");
                        return [4 /*yield*/, this._spHttpClient.get(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        return [4 /*yield*/, response.json()];
                    case 2:
                        json = _a.sent();
                        if (!(json === null || json === void 0 ? void 0 : json.value)) return [3 /*break*/, 7];
                        allItems = json.value;
                        uniqueItems = allItems.filter(function (i) { return i.HasUniqueRoleAssignments === true; });
                        if (uniqueItems.length === 0)
                            return [2 /*return*/, []];
                        results_1 = [];
                        chunks = this.chunkArray(uniqueItems, 5);
                        _i = 0, chunks_1 = chunks;
                        _a.label = 3;
                    case 3:
                        if (!(_i < chunks_1.length)) return [3 /*break*/, 6];
                        chunk = chunks_1[_i];
                        promises = chunk.map(function (item) { return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(_this, void 0, void 0, function () {
                            var permEndpoint, permResp, permJson, roles, e_1;
                            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        _a.trys.push([0, 3, , 4]);
                                        permEndpoint = "".concat(this._webUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(item.Id, ")/roleassignments?$expand=Member,RoleDefinitionBindings");
                                        return [4 /*yield*/, this._spHttpClient.get(permEndpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1)];
                                    case 1:
                                        permResp = _a.sent();
                                        return [4 /*yield*/, permResp.json()];
                                    case 2:
                                        permJson = _a.sent();
                                        roles = (permJson === null || permJson === void 0 ? void 0 : permJson.value) ? this._processRoleAssignments(permJson.value) : [];
                                        return [2 /*return*/, {
                                                Id: item.Id,
                                                Title: item.FileLeafRef || item.Title, // Use filename for files
                                                ServerRelativeUrl: item.FileRef,
                                                FileSystemObjectType: item.FileSystemObjectType,
                                                RoleAssignments: roles
                                            }];
                                    case 3:
                                        e_1 = _a.sent();
                                        console.error("Error fetching perms for item ".concat(item.Id), e_1);
                                        return [2 /*return*/, null];
                                    case 4: return [2 /*return*/];
                                }
                            });
                        }); });
                        return [4 /*yield*/, Promise.all(promises)];
                    case 4:
                        chunkResults = _a.sent();
                        chunkResults.forEach(function (r) { if (r)
                            results_1.push(r); });
                        _a.label = 5;
                    case 5:
                        _i++;
                        return [3 /*break*/, 3];
                    case 6: return [2 /*return*/, results_1];
                    case 7: return [2 /*return*/, []];
                    case 8:
                        error_6 = _a.sent();
                        console.error("Error scanning items for list ".concat(listId), error_6);
                        return [2 /*return*/, []];
                    case 9: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype.chunkArray = function (myArray, chunk_size) {
        var results = [];
        var arrayCopy = (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__spreadArray)([], myArray, true);
        while (arrayCopy.length) {
            results.push(arrayCopy.splice(0, chunk_size));
        }
        return results;
    };
    PermissionServiceImpl.prototype.removeSitePermission = function (principalId) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                endpoint = "".concat(this._webUrl, "/_api/web/roleassignments/getbyprincipalid(").concat(principalId, ")");
                return [2 /*return*/, this._removeRoleAssignment(endpoint)];
            });
        });
    };
    PermissionServiceImpl.prototype.removeListPermission = function (listId, principalId) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                endpoint = "".concat(this._webUrl, "/_api/web/lists(guid'").concat(listId, "')/roleassignments/getbyprincipalid(").concat(principalId, ")");
                return [2 /*return*/, this._removeRoleAssignment(endpoint)];
            });
        });
    };
    PermissionServiceImpl.prototype.removeItemPermission = function (listId, itemId, principalId) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var endpoint;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                endpoint = "".concat(this._webUrl, "/_api/web/lists(guid'").concat(listId, "')/items(").concat(itemId, ")/roleassignments/getbyprincipalid(").concat(principalId, ")");
                return [2 /*return*/, this._removeRoleAssignment(endpoint)];
            });
        });
    };
    PermissionServiceImpl.prototype._removeRoleAssignment = function (endpoint) {
        return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__awaiter)(this, void 0, void 0, function () {
            var response, error_7;
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__generator)(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._spHttpClient.post(endpoint, _microsoft_sp_http__WEBPACK_IMPORTED_MODULE_0__.SPHttpClient.configurations.v1, {
                                headers: {
                                    'Accept': 'application/json',
                                    'Content-Type': 'application/json',
                                    'X-HTTP-Method': 'DELETE',
                                    'IF-MATCH': '*'
                                }
                            })];
                    case 1:
                        response = _a.sent();
                        return [2 /*return*/, response.ok];
                    case 2:
                        error_7 = _a.sent();
                        console.error("Error removing permission", error_7);
                        return [2 /*return*/, false];
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    PermissionServiceImpl.prototype._processRoleAssignments = function (assignments) {
        if (!assignments)
            return [];
        return assignments.map(function (assignment) {
            // Filter out "Limited Access" from bindings
            var filteredBindings = assignment.RoleDefinitionBindings.filter(function (binding) { return binding.Name !== 'Limited Access'; });
            return (0,tslib__WEBPACK_IMPORTED_MODULE_1__.__assign)((0,tslib__WEBPACK_IMPORTED_MODULE_1__.__assign)({}, assignment), { RoleDefinitionBindings: filteredBindings });
        }).filter(function (assignment) { return assignment.RoleDefinitionBindings.length > 0; });
    };
    return PermissionServiceImpl;
}());



/***/ }),

/***/ 91909:
/*!*************************************!*\
  !*** external "@microsoft/sp-http" ***!
  \*************************************/
/***/ ((module) => {

module.exports = __WEBPACK_EXTERNAL_MODULE__91909__;

/***/ })

},
/******/ function(__webpack_require__) { // webpackRuntimeModules
/******/ /* webpack/runtime/getFullHash */
/******/ (() => {
/******/ 	__webpack_require__.h = () => ("f6f87640fd2a2c022a6e")
/******/ })();
/******/ 
/******/ }
);
//# sourceMappingURL=permission-viewer-web-part.ca2d52bd176eecd06765.hot-update.js.map