// ==UserScript==
// @name         Teams API Interaction
// @match        https://teams.microsoft.com.mcas.ms/v2/*
// @description  Teams API Interaction
// @version      2026-04-14
// @grant       unsafeWindow
// ==/UserScript==

(function () {
    'use strict';

    const API = {
        meeting: {
            check: {
                doesMeetingMiniatureExist() {
                    const miniature = document.querySelector('[data-app-name="floatingmonitor"]');
                    if (miniature && miniature.offsetHeight > 0) {
                        return 'yes';
                    }
                    return 'no';
                },
                isOnFullScreenMeeting() {
                    if (this.doesMeetingMiniatureExist() === 'no') {
                        const hangupButton = document.querySelector('#hangup-button');
                        return hangupButton ? 'yes' : 'no';
                    }
                    return 'no';
                }
            },
            fullScreen: {
                toggleChatWindow() {
                    const chatButton = document.querySelector('[data-inp="chat-button"]');
                    if (chatButton) {
                        chatButton.click();
                    }
                },
                toggleRosterWindow() {
                    const rosterButton = document.querySelector('[data-inp="roster-button"]');
                    if (rosterButton) {
                        rosterButton.click();
                    }
                },
                raiseHand: {
                    isHandRaised() {
                        const raiseHandButton = document.querySelector('[data-inp="raisehands-button"]');
                        if (raiseHandButton) {
                            const state = raiseHandButton.getAttribute('data-track-module-name');
                            if (state === 'raiseHandButton') {
                                return 'no';
                            } else if (state === 'lowerHandButton') {
                                return 'yes';
                            }
                        }
                    },
                    toggleRaiseHand() {
                        const raiseHandButton = document.querySelector('[data-inp="raisehands-button"]');
                        if (raiseHandButton) {
                            raiseHandButton.click();
                        }
                    },
                },
                react: {
                    toggleActionsMenu() {
                        const reactionsButton = document.querySelector('[data-inp="reaction-menu-button"]');
                        if (reactionsButton) {
                            reactionsButton.click();
                        }
                    },
                    isReactionsMenuOpen() {
                        const reactionsMenu = document.querySelector('[data-tid="reaction-menu-button-toolbox"]');
                        return reactionsMenu ? 'yes' : 'no';
                    },
                    reaction: {
                        like() {
                            const likeButton = document.querySelector('[data-inp="like-button"]');
                            if (likeButton) {
                                likeButton.click();
                            }
                        },
                        love() {
                            const loveButton = document.querySelector('[data-inp="heart-button"]');
                            if (loveButton) {
                                loveButton.click();
                            }
                        },
                        applause() {
                            const applauseButton = document.querySelector('[data-inp="applause-button"]');
                            if (applauseButton) {
                                applauseButton.click();
                            }
                        },
                        laugh() {
                            const laughButton = document.querySelector('[data-inp="laugh-button"]');
                            if (laughButton) {
                                laughButton.click();
                            }
                        },
                        surprise() {
                            const surpriseButton = document.querySelector('[data-inp="surprised-button"]');
                            if (surpriseButton) {
                                surpriseButton.click();
                            }
                        }
                    }
                }
            },
            miniature: {
                expand() {
                    const expandButton = document.querySelector('[data-tid="call-monitor-navigate-cw-button"]');
                    if (expandButton) {
                        expandButton.click();
                    }
                }
            },
            actions: {
                camera: {
                    toggle() {
                        const cameraButton = document.querySelector('[data-inp="video-button"]');
                        if (cameraButton) {
                            cameraButton.click();
                        }
                    },
                    getStatus() {
                        const cameraButton = document.querySelector('[data-inp="video-button"]');
                        if (cameraButton) {
                            const state = cameraButton.getAttribute('data-state');
                            if (state === 'call-video-off') {
                                return 'off';
                            } else if (state === 'call-video') {
                                return 'on';
                            }
                        }
                        return 'unknown';
                    }
                },
                microphone: {
                    toggle() {
                        const micButton = document.querySelector('[data-inp="microphone-button"]');
                        if (micButton) {
                            micButton.click();
                        }
                    },
                    getStatus() {
                        const micButton = document.querySelector('[data-inp="microphone-button"]');
                        if (micButton) {
                            const state = micButton.getAttribute('data-state');
                            if (state === 'mic-off') {
                                return 'off';
                            } else if (state === 'mic') {
                                return 'on';
                            }
                        }
                        return 'unknown';
                    }
                },
                shareScreen: {
                    toggle: async function () {
                        const shareButton = document.querySelector('[data-inp="share-button"]');
                        const status = this.getStatus();

                        if (shareButton) {
                            shareButton.click();
                        }

                        if (status === 'off') {
                            await new Promise(r => setTimeout(r, 500));

                            const screenShareOption = document.querySelector('[data-tid="share-screen-window-or-tab"]');
                            if (screenShareOption) {
                                screenShareOption.click();
                            }
                        }
                    },
                    getStatus() {
                        // on: call-control-stop-presenting-new
                        // off: call-control-present-new
                        const shareButton = document.querySelector('[data-inp="share-button"]');
                        if (shareButton) {
                            const state = shareButton.getAttribute('data-state');
                            if (state === 'call-control-present-new') {
                                return 'off';
                            } else if (state === 'call-control-stop-presenting-new') {
                                return 'on';
                            }
                        }
                        return 'unknown';
                    }
                },
                hangup() {
                    const hangupButton = document.querySelector('[data-inp="hangup-button"]');
                    if (hangupButton) {
                        hangupButton.click();
                    }
                }
            }
        },
        general: {
            chats: {
                goToChats() {
                    const chatsButton = document.querySelector('button[aria-label^="Conversation"]:not([data-inp="chat-button"])');
                    if (chatsButton) {
                        chatsButton.click();
                    }
                },
                getChatsNames(){
                    const chats = document
                    .querySelector('[data-item-type="chats"]')
                    ?.querySelectorAll('[data-item-type="chat"]');

                    const texts = [...chats].flatMap(chat =>
                    [...chat.querySelectorAll('[id^="title-chat-list-item"]')]
                        .map(el => el.innerText)
                    );
                    return texts;
                },
                clickOnChatByName(name) {
                    const chats = document
                    .querySelector('[data-item-type="chats"]')
                    ?.querySelectorAll('[data-item-type="chat"]');

                    for (const chat of chats) {
                        const titleElement = chat.querySelector('[id^="title-chat-list-item"]');
                        if (titleElement && titleElement.innerText === name) {
                            titleElement.click();
                            break;
                        }
                    }
                }
            },
            goToCalendar() {
                const calendarButton = document.querySelector('button[aria-label^="Calendrier"]');
                if (calendarButton) {
                    calendarButton.click();
                }
            },
            goToActivity() {
                const activityButton = document.querySelector('button[aria-label^="Activité"]');
                if (activityButton) {
                    activityButton.click();
                }
            }
        }
    };


    unsafeWindow.myTeamsAPIName = API;

})();